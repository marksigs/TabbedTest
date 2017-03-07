VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 2  'RequiresTransaction
END
Attribute VB_Name = "ApplicationManagerTxBO"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
'Workfile:      ApplicationManagerTxBO.cls
'Copyright:     Copyright © 1999 Marlborough Stirling

'Description:   Code template for omiga4 Business Object which supports transactions
'               contains (implied) standard interface:
'               Public Sub Create(vstrXMLRequest As String) As String
'               Public Sub Update(vstrXMLRequest As String) As String
'               Public Sub Delete(vstrXMLRequest As String) As String

'Dependencies:  ApplicationManagerDO
'Issues:        not part of public interface
'               Instancing:         MultiUse
'               MTSTransactionMode: RequiresTransactions
'------------------------------------------------------------------------------------------
'History:
'
' Prog  Date     Description
' SR    12/01/00 Restructured
' SR    28/03/00 AQR/SYS0015 - Modified method - 'LockCustomersForApplication'
'                Raise proper error when a duplicate record is found during creation of
'                Application, Customer, CustomerApplicationlocks
'SR     29/03/00 New function 'GetUserHoldingCustomerOrApplication'.Modified method
'                'LockCustomersForApplication' to display the user and Application details
'                in the message, when duplicate records are found
'IW     23/05/00 SYS0774 - DISTRIBUTIONCHANNELID S/B CHANNELID
'MC     06/06/00 SYS0611 - Added function AddLockTypeAndDate to ensure LockType and Date
'                          are saved when locking an existing application.
'MC     20/06/00 Amend CreateFullApplicationDetails for additional Import data
'JLD    27/06/00 Do not create a CustomerLockApplicationLock if the global 'Locking' parameter is '0'
'SR     13/09/00 SYS1138, SYS1435 New Method SaveCustomerAndEmploymentDetails
'LD     07/11/00 Explicity close database connections
'CL     30/11/00 CORE000006 Amend to use new IAuditDO interface
'PSC    12/01/01 SYS1749 Add ImportAccountsIntoApplication()
'PSC    22/02/01 SYS1749 Amend format of returned data to have application tag
'PSC    23/02/01 SYS1749 Amend ImportAccountAndLoans to take into account no existing loans
'PSC    26/02/01 SYS1749 Amend ImportCustomers to clear out the customer name
'ASt    15/03/01 CC012 Compatible with SQL Server and Oracle
'JLD    23/10/01 SYS2839 Check input XML for LOCKING node.
'SG     04/04/02 SYS4360 Amended ImportCustomers to correct CustomerRole creation code.
'SG     12/04/02 SYS4384 Amended ImportArrearsHistory, validate ArrearsHistory xml before db update
'STB    29/04/02 SYS4508 Enable CustomerRole linking for primary downloaded customer.
'MEVA   29/04/02 SYS1452 Do not report missing locks
'SG     22/05/02 SYS4638 Amended IApplicationManagerTxBO_ImportAccountsIntoApplication
'STB    27/05/02 SYS4637 NewProperty details not populated when importing a further advance.
'DS     28/05/02 SYS4703 ARREARSHISTORY is not mandatory.
'DRC    11/06/02 SYS4740 In ImportAccountsintoApplications, with a Further Advance
'                        the othersystemaccountnumber is set to the existing adminsystem account number
'------------------------------------------------------------------------------------------

Option Explicit
Implements IApplicationManagerTxBO
Implements ObjectControl
 
'=============================================
'Constant Declaration Section
'=============================================
Private Const cstrROOT_NODE_NAME = "APPLICATION"
Private Const cstrTABLE_CUSTLOCKAPPLLOCK = "CUSTOMERLOCKAPPLICATIONLOCK"
Private Const cstrTABLE_APPLLOCK = "APPLICATIONLOCK"
Private Const cstTABLE_CUSTOMERROLE = "CUSTOMERROLE"
Private Const cstTABLE_CUSTOMER = "CUSTOMER"
Private Const cstrCOMPONENT_NAME = "omApp"
Private Const cstrBASE = "omBase"
Private Const cstrGetAddData = "GETADDDATAFORACCDOWNLOAD"  ' SR 06/08/2004 : BMIDS830
'=============================================
'Variable Declaration Section
'=============================================

Private m_objContext As ObjectContext
Private m_objErrAssist As ErrAssist
Private m_objXmlAssist As XMLAssist             ' Remove if not required


'------------------------------------------------------------------------------------------
' BMids History:
'
' Prog  Date        Description
' MDC   21/05/2002  BMIDS00004 - BM073 Versions of Employment
' MV    20/06/2002  BMIDS0077 - Core Upgrade 7.0.2 - Modified ImportCustomers()
' MDC   18/07/2002  BMIDS00006 - BM051/BM054 Customer Account Download
' GHun  08/08/2002  BMIDS00006 - BM051/BM054 Customer Account Download
' GHun  06/09/2002  BMIDS00425 - Changed ImportAccountsIntoApplication, ImportPropertyAddress, CreateMortgageAccount, ImportArrearsHistory
' GHun  17/09/2002  BMIDS00475 - Format GetAccountCustomers XML as per DTD before passing to admin system
' GHun  03/10/2002  BMIDS00566 - Apply core AQR SYS4898 to update correspondence salutation
' GHun  09/20/2002  BMIDS00508 - Set BMids mortgage indicator in CreateMortgageAccount
' PSC   16/10/2002  BMIDS00575 - Changed ImportAccountsIntoApplication to set up current
'                                mortgage account list correctly
' PSC   23/10/2002  BMIDS00465 - Amend ImportAccountsIntoApplication to only get accounts from the
'                                admin system if the customer has an OTHERSYSTEMCUSTOMERNUMBER
' MDC   24/10/2002  BMIDS00527 - Update ApplicationFactFind with NumberOfApplicants and NumberOfGuarantors
'                                when creating an application during Account Download.
' MDC   24/10/2002  BMIDS00682 - Do not create Additional Current CustomerAddress in Account Download
' GHun  29/10/2002  BMIDS00701 - ConvertCustomerToOmigaFormat: old Omiga3 customer numbers should be used if available
' GHun  30/10/2002  BMIDS00447 - ImportArrearsHistory: default description of loan if it does not exist and change validation
' GHun  18/11/2002  BMIDS00758 - Amend ImportPropertyAddress to also match on street
' GHun  18/11/2002  BMIDS00563 - Amend ImportAccountsIntoApplication so reimporting an
'                                account does not create duplicate NewProperty records
' SR    20/11/2002  BMIDS00500   [Core SYS5040] Altered LockCustomersForApplication, UnlockApplicationAndCustomers
'                                to use stored procs.
' GHun  20/11/2002  BMIDS01026 - Amend MaintainCustomerRoles to update further advance rule
' GHun  22/11/2002  BMIDS01053 - Amend ImportAccountsIntoApplication to fix No Records Found handling
' GHun  26/11/2002  BMIDS01043 - Amend ImportAccountsIntoApplication so OtherSystemAccountNumber is updated in a
'                                new transaction before anything else can fail
' GHun  08/01/2003  BM0233     - Amend ImportAccountsIntoApplication to save PropertyTenure correctly
' GHun  25/02/2003  BM0377     - Amend ImportAccountsIntoApplication to not overwrite FinancialSummary
' MV    14/03/2003  BM0358     - Amended IApplicationManagerTxBO_ImportAccountsIntoApplication() to Save SpecialFeature Data
' GHun  17/04/2003  BM0308     - Amend ConvertCustomerToOmigaFormat to handle blank dates only containing slashes
' GHun  17/04/2003  BM0387     - Amend ImportAccountsIntoApplication to correctly set ArrearsHistoryIndicator
' GHun  22/04/2003  BM0465     - Amend ImportPropertyAddress to support blank postcodes
' GD    30/07/2003  BMIDS614     Amend CreateMortgageAccount and UpdateMortgageAccount to use MONTHLYRENTALINCOME.
' PSC   03/10/2003  BMIDS612   - Amend ImportCustomers to use OmigaCustomerNumber to search for legacy customer
' RF    06/10/03    BMIDS642     Fix problems with apostrophes within application / existing account data.
'                                Improve error handling in ImportPropertyAddress.
' PSC   03/10/2003  BMIDS655   - Amend ImportCustomers to make OmigaCustomerNumber optional
' DRC   04/12/2003  bmids667     Amend CreateMortgageAccount and UpdateMortgageAccount to use LASTVALUERNAME.
' DRC   25/02/2004  BMIDS718    Amend ImportAccountsIntoApplication for GetAccountDetails DTD changes
' KRW   18/05/2004  BMIDS745    Amended ConvertCustomerToOmigaFormat to Capture Nationality for existing ICBS Customers
' HMA   13/07/2004  BMIDS756    Amended CreateMortgageAccount
' GHun  14/07/2004  BMIDS756    Amended CreateMortgageAccount and minor tidy up
' INR   15/07/04    BMIDS756    Reformat the sortcode if required
' GHun  15/07/2004  BMIDS756    Amended UpdateMortgageLoan and ImportAccountAndLoans
' HMA   21/07/2004  BMIDS758    Update RemovedToECustomer table on delete
' JD    03/08/2004  BMIDS830    CC079 add more information to the GetAccountDetails request.
' HMA   04/08/2004  BMIDS836    Change validation type to "CC" for CC069
' RF    05/08/2004  BMIDS818    Performance of ImportAccountsIntoApplication
' SR    09/08/2004  BMIDS830    Modified 'ImportAccountsIntoApplication'
' JD    10/08/2004  BMIDS830    Change tags in 'ImportAccountsIntoApplication'
' SR    13/08/2004  BMIDS815    Modified method 'ImportAccountsIntoApplication' - Call to ApplicationDO.Update
'                               to update LastAccountDowloadDate
' SR    23/08/2004  BMIDS815    New method - 'ImportAccountRefreshIntoApplication'
' SR    03/09/2004  BMIDS815
' HMA   09/09/2004  BMIDS864    Save Customer Verification
' GHun  01/10/2004  BMIDS907    Changed UpdateMortgageLoan to save new porting fields
' GHun  22/10/2004  BMIDS918    Performance improvements for ImportAccountsIntoApplication
' JD    16/03/2005  BMIDS975    Update term on mortgageloan in ImportAccountRefreshIntoApplication.
'-----------------------------------------------------------------------------------------------------------------
' MARS History:
'
' Prog  Date        Description
' HMA   11/07/2005  MAR11       Apply BMIDS989 Change ImportPropertyAddress
' SR    01/05/2006  MAR1631     Create AccessAudit records on locking and unlocking of application
'                               modified methods 'LockCustomersForApplication',UnlockApplicationAndCustomers
'----------------------------------------------------------------------------------------------------------------
' Epsom History:
'
' Prog  Date        Description
' GHun  18/12/2006  EP2_56      Changes for TOE
' SR    15/01/2007  EP2_858     modified private functions 'CreateMortgageAccount', 'UpdateMortgageAccount',
'                               'ImportAccountAndLoans'
'                               and IApplicationManagerTxBO_ImportAccountsIntoApplication
' SR    23/01/2007  EP2_858
' SR    24/01/2007  EP2_858     modified IApplicationManagerTxBO_ImportAccountsIntoApplication.
' PSC   15/02/2007  EP2_858     Add addition fields to CreateMortgageAccount and UpdateMortgageAccount
' PSC   30/01/2007  EP2_1114    Correct creation of application bank building society
' PSC   30/01/2007  EP2_1146    Create bank building society if product switch, TOE or port
'                               Update application with account guid if additional borrowing, product switch, TOE or port
'PSC    01/02/2007  EP2_1113    Amend ImportAccountAndLoans to set loan redemption status correctly
'PSC    07/02/2007  EP2_1260    Amend ImportAccountAndLoans to set up product start date correctly
'PSC    07/02/2007  EP2_1271    Amend ImportAccountAndLoans to not set up product start date
'PSC    16/02/2007  EP2_1314    Amend ImportAccountAndLoans to not use ABO validation type
'MAH    13/03/2007  EP2_1597    Supress populating new property details with old property details when porting a loan
'PSC    30/03/2007  EP2_1636    Change redemption status for switch to from PS to PSR
'----------------------------------------------------------------------------------------------------------------

Private Sub CreateCustomerRole(ByVal vxmlRequest As IXMLDOMElement)
' header ----------------------------------------------------------------------------------
' description:
'   Create a CustomerRole.
' pass:
'   vstrXMLRequest  Format: As per AddCustomerToApplication
' return:   n/a
' Raise Errors:
'------------------------------------------------------------------------------------------
On Error GoTo CreateCustomerRoleVbErr

    Const strFunctionName As String = "CreateCustomerRole"

    Dim objRequestDoc As FreeThreadedDOMDocument40
    Dim objXmlElem As IXMLDOMElement
    Dim objIApplicationDO As IApplicationDO
    Dim objIADOAssist As IADOAssist
    Dim objIApplManagerDO As IApplicationManagerDO
    Dim objGlobalParamDO As GlobalParameterDO

    Dim objSqlssist As New SQLAssist

    Dim objXmlTempNode As IXMLDOMNode
    Dim objTableElem As IXMLDOMElement
    Dim objXmlRequestNode As IXMLDOMNode

    Dim strCustomerRoleType As String
    Dim strCustomerOrder As String
    Dim strApplicationNumber As String, strAFFNumber As String
    Dim strCustomerNumber As String
    
    strCustomerRoleType = _
        m_objXmlAssist.GetTagValue(vxmlRequest, "CUSTOMERROLETYPE")
    strApplicationNumber = m_objXmlAssist.GetTagValue(vxmlRequest, "APPLICATIONNUMBER")

    strAFFNumber = m_objXmlAssist.GetTagValue( _
                                vxmlRequest, "APPLICATIONFACTFINDNUMBER")

    If Len(strCustomerRoleType) = 0 Then
    
        If m_objContext Is Nothing Then
            Set objIApplicationDO = New ApplicationDO
            Set objIADOAssist = New ADOAssist
        Else
            Set objIApplicationDO = _
                    m_objContext.CreateInstance(cstrCOMPONENT_NAME & ".ApplicationDO")
            Set objIADOAssist = _
                    m_objContext.CreateInstance(gstrBASE_COMPONENT & ".ADOAssist")
        End If

        ' set CustomerRole to Applicant, i.e. type 1
        strCustomerRoleType = "1"

        ' set CustomerOrder to be 1 more than the NumberOfApplicants for the ApplicationFactFind
        Dim strNumberOfApplicants As String
        
        Dim strCondition As String
        strCondition = " APPLICATIONNUMBER = " & objSqlssist.FormatString(strApplicationNumber) & " AND" _
                       & " APPLICATIONFACTFINDNUMBER = " & strAFFNumber

        If IsNull(objIADOAssist.GetValueFromTable("APPLICATIONFACTFIND", strCondition, "NUMBEROFAPPLICANTS", _
                                           strNumberOfApplicants)) Then
            m_objErrAssist.ThrowError TypeName(Me), strFunctionName, _
                oeInternalError, "NUMBEROFAPPLICANTS not set"
        End If

        strCustomerOrder = CStr(CInt(strNumberOfApplicants) + 1)
    Else
        ' leave CustomerRole as its input value;
        ' set CustomerOrder to its input value
        strCustomerOrder = _
            m_objXmlAssist.GetTagValue(vxmlRequest, "CUSTOMERORDER")

        If Len(strCustomerOrder) < 1 Then
            m_objErrAssist.ThrowError TypeName(Me), strFunctionName, _
                 oeInvalidParameter, _
                "If CUSTOMERROLETYPE is specified, CUSTOMERORDER must also be specified"
        End If
    End If

    ' add the <REQUEST> node
    Set objRequestDoc = New FreeThreadedDOMDocument40
    objRequestDoc.validateOnParse = False
    objRequestDoc.setProperty "NewParser", True
    Set objXmlRequestNode = m_objXmlAssist.GetRequestNode(vxmlRequest.ownerDocument)
    objRequestDoc.appendChild objXmlRequestNode

    ' add the table name
    Set objXmlElem = vxmlRequest.ownerDocument.createElement("CUSTOMERROLE")
    Set objTableElem = objXmlRequestNode.appendChild(objXmlElem)

    ' add the CustomerRole fields
    Dim strTagName As String

    strTagName = "CUSTOMERNUMBER"
    Set objXmlElem = vxmlRequest.ownerDocument.createElement(strTagName)
    strCustomerNumber = m_objXmlAssist.GetTagValue(vxmlRequest, strTagName)
    objXmlElem.Text = strCustomerNumber
    objTableElem.appendChild objXmlElem

    strTagName = "CUSTOMERVERSIONNUMBER"
    Set objXmlElem = vxmlRequest.ownerDocument.createElement(strTagName)
    objXmlElem.Text = m_objXmlAssist.GetTagValue(vxmlRequest, strTagName)
    objTableElem.appendChild objXmlElem

    Set objXmlElem = vxmlRequest.ownerDocument.createElement("CUSTOMERROLETYPE")
    objXmlElem.Text = strCustomerRoleType
    objTableElem.appendChild objXmlElem

    Set objXmlElem = vxmlRequest.ownerDocument.createElement("CUSTOMERORDER")
    objXmlElem.Text = strCustomerOrder
    objTableElem.appendChild objXmlElem

    strTagName = "APPLICATIONNUMBER"
    Set objXmlElem = vxmlRequest.ownerDocument.createElement(strTagName)
    objXmlElem.Text = strApplicationNumber
    objTableElem.appendChild objXmlElem

    strTagName = "APPLICATIONFACTFINDNUMBER"
    Set objXmlElem = vxmlRequest.ownerDocument.createElement(strTagName)
    objXmlElem.Text = strAFFNumber
    objTableElem.appendChild objXmlElem

    '   Create CustomerRole record
    If m_objContext Is Nothing Then
        Set objIApplManagerDO = New ApplicationManagerDO
        Set objGlobalParamDO = New GlobalParameterDO
    Else
        Set objIApplManagerDO = m_objContext.CreateInstance(cstrCOMPONENT_NAME & _
                                                         ".ApplicationManagerDO")
        Set objGlobalParamDO = m_objContext.CreateInstance(cstrBASE & ".GlobalParameterDO")
    End If

     objIApplManagerDO.CreateCustomerRole objXmlRequestNode
    
    '   Create Lock (record in table 'CustomerLockApplicationLock') if required
    Dim strLockingReqd As String
    'JLD Check input XML for LOCKING node first
    strLockingReqd = m_objXmlAssist.GetTagValue(vxmlRequest, "LOCKING", , True)
    If strLockingReqd = "" Then 'no locking node found in input XML so look in the database
        strLockingReqd = objGlobalParamDO.GetCurrentParameterByType("Locking", "BOOLEAN")
    End If
    If strLockingReqd = "1" Then
        Set objRequestDoc = New FreeThreadedDOMDocument40
        objRequestDoc.validateOnParse = False
        objRequestDoc.setProperty "NewParser", True
        Set objXmlRequestNode = m_objXmlAssist.GetRequestNode(vxmlRequest.ownerDocument)
        objRequestDoc.appendChild objXmlRequestNode
        Set objTableElem = objRequestDoc.createElement("CUSTOMERLOCKAPPLICATIONLOCK")
        objXmlRequestNode.appendChild objTableElem
        Set objXmlElem = objRequestDoc.createElement("APPLICATIONNUMBER")
        objXmlElem.Text = strApplicationNumber
        objTableElem.appendChild objXmlElem
        Set objXmlElem = objRequestDoc.createElement("CUSTOMERNUMBER")
        objXmlElem.Text = strCustomerNumber
        objTableElem.appendChild objXmlElem
          
        objIApplManagerDO.CreateCustomerLockApplicationLock objXmlRequestNode
    End If
    
    If Not m_objContext Is Nothing Then
        m_objContext.SetComplete
    End If
    
CreateCustomerRoleExit:
    
    Set objXmlElem = Nothing
    Set objXmlTempNode = Nothing
    Set objXmlRequestNode = Nothing
    Set objRequestDoc = Nothing
    Set objIApplicationDO = Nothing
    Set objIADOAssist = Nothing
    Set objSqlssist = Nothing
    Set objIApplManagerDO = Nothing
    
    Exit Sub
CreateCustomerRoleVbErr:
    Set objXmlElem = Nothing
    Set objXmlTempNode = Nothing
    Set objXmlRequestNode = Nothing
    Set objRequestDoc = Nothing
    Set objIApplicationDO = Nothing
    Set objIADOAssist = Nothing
    Set objSqlssist = Nothing
    Set objIApplManagerDO = Nothing

    If m_objErrAssist.IsSystemError Then
        m_objErrAssist.LogError TypeName(Me), strFunctionName, Err.Number, Err.Description
    End If
    
    m_objErrAssist.AddToErrSource strFunctionName

    If Err.Number = oeDuplicateKey Then
        ' "Customer already linked to application"
        m_objErrAssist.ThrowError TypeName(Me), strFunctionName, 139
    End If
        
    ' Re-raise error for business object to interpret as appropriate
    Err.Raise Err.Number, Err.Source, Err.Description
    
End Sub

Private Function IApplicationManagerTxBO_AddCustomerToApplication( _
                        ByVal vxmlRequest As IXMLDOMElement) As IXMLDOMNode
' header ----------------------------------------------------------------------------------
' description:
'   Link a customer to an application. In all cases both the customer and the
'   application records will already exist.
' pass:
'   vXMLRequest  Format:
'                   <REQUEST>
'                       <CUSTOMERNUMBER>
'                       </CUSTOMERNUMBER>
'                       <CUSTOMERVERSIONNUMBER>
'                       </CUSTOMERVERSIONNUMBER>
'                       <CUSTOMERROLETYPE>
'                       </CUSTOMERROLETYPE>
'                       <CUSTOMERORDER>
'                       </CUSTOMERORDER>
'                       <APPLICATIONNUMBER>
'                       </APPLICATIONNUMBER>
'                       <APPLICATIONFACTFINDNUMBER>
'                       </APPLICATIONFACTFINDNUMBER>
'                   </REQUEST>
' return:   n/a
' Raise Errors:
'------------------------------------------------------------------------------------------
On Error GoTo AddCustomerToApplicationVbErr

    Const strFunctionName As String = "IApplicationManagerTxBO_AddCustomerToApplication"
    
    Dim xmlOut As FreeThreadedDOMDocument40
    Dim xmlResponseElem As IXMLDOMElement
    Dim xmlTempResponseNode As IXMLDOMNode
    
    Set xmlOut = New FreeThreadedDOMDocument40
    xmlOut.validateOnParse = False
    xmlOut.setProperty "NewParser", True
    Set xmlResponseElem = xmlOut.createElement("RESPONSE")
    xmlOut.appendChild xmlResponseElem
    xmlResponseElem.setAttribute "TYPE", "SUCCESS"
    
     '--------------------------------------------
    ' create CustomerRole
    '---------------------------------------------
    CreateCustomerRole vxmlRequest

    '----------------------------------------------
    ' validate the change
    '----------------------------------------------
     Call MaintainAddCustomerToApplication(vxmlRequest)
   
    Set IApplicationManagerTxBO_AddCustomerToApplication = xmlResponseElem
    
    If Not m_objContext Is Nothing Then
        m_objContext.SetComplete
    End If
    
AddCustomerToApplicationExit:
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

    Set IApplicationManagerTxBO_AddCustomerToApplication = _
                                m_objErrAssist.CreateErrorResponseEx()

    If Not m_objContext Is Nothing Then
        m_objContext.SetAbort
    End If
    
    ' go to clean-up section
    Resume AddCustomerToApplicationExit

End Function

Private Sub MaintainAddCustomerToApplication(ByVal vxmlRequest As IXMLDOMElement)
'----------------------------------------------------------------------------------
'Description : This function performs all the updations to maintain the customer
'              roles after adding customer to an application.
'              N.B. A seperate function is written because of the performance
'              problems.
'----------------------------------------------------------------------------------
On Error GoTo MaintainAddCustomerToApplicationVbErr

    Const strFunctionName As String = "MaintainAddCustomerToApplication"

    Dim xmlInputElem As IXMLDOMElement 'Input to ApplicationFactFind
    Dim xmlElem As IXMLDOMElement, xmlTableElem As IXMLDOMElement
    Dim xmlTempResponseNode As IXMLDOMElement, xmlResponseElem As IXMLDOMElement
    
    Dim strApplCount As String, strApplicationNumber As String
    Dim intAFFNumber As Integer
    
    Dim objIApplicationDO As IApplicationDO
    Dim objIADOAssist As IADOAssist
    'EP2_56 GHun
    Dim strApplicationType As String
    Dim objComboDO As ComboDO
    Dim xmlOut  As FreeThreadedDOMDocument40
    Dim strTag As String
    Dim objIApplManagerDO As IApplicationManagerDO
    'EP2_56 End
    
    Set xmlResponseElem = vxmlRequest.ownerDocument.createElement("RESPONSE")
    xmlResponseElem.setAttribute "TYPE", "SUCCESS"
    
    '   Build Request node with either 'NumberOfApplicants' or 'NumberOfGuarantors'
    '   based on which ever is to be updated (not empty in the request)
    Set xmlInputElem = m_objXmlAssist.GetRequestNode(vxmlRequest.ownerDocument)
    Set xmlTableElem = vxmlRequest.ownerDocument.createElement("APPLICATIONFACTFIND")
    xmlInputElem.appendChild xmlTableElem
    
    Set xmlElem = vxmlRequest.ownerDocument.createElement("APPLICATIONNUMBER")
    strApplicationNumber = vxmlRequest.selectSingleNode(".//" & "APPLICATIONNUMBER").Text
    xmlElem.Text = strApplicationNumber
    xmlTableElem.appendChild xmlElem
    
    Set xmlElem = vxmlRequest.ownerDocument.createElement("APPLICATIONFACTFINDNUMBER")
    intAFFNumber = vxmlRequest.selectSingleNode(".//" & "APPLICATIONFACTFINDNUMBER").Text
    xmlElem.Text = intAFFNumber
    xmlTableElem.appendChild xmlElem
    
    strApplCount = m_objXmlAssist.GetTagValue(vxmlRequest, "NUMBEROFAPPLICANTS")
    If strApplCount <> "" Then
        Set xmlElem = vxmlRequest.ownerDocument.createElement("NUMBEROFAPPLICANTS")
        xmlElem.Text = strApplCount
        xmlTableElem.appendChild xmlElem
    Else
        strApplCount = m_objXmlAssist.GetTagValue(vxmlRequest, "NUMBEROFGUARANTORS")
        Set xmlElem = vxmlRequest.ownerDocument.createElement("NUMBEROFGUARANTORS")
        xmlElem.Text = strApplCount
        xmlTableElem.appendChild xmlElem
    End If
    
    If m_objContext Is Nothing Then
        Set objIApplicationDO = New ApplicationDO
        Set objIADOAssist = New ADOAssist
    Else
        Set objIApplicationDO = m_objContext.CreateInstance(App.Title & ".ApplicationDO")
        Set objIADOAssist = m_objContext.CreateInstance(gstrBASE_COMPONENT & ".ADOAssist")
    End If
    
    objIApplicationDO.UpdateApplicationFactFind xmlInputElem
     
    '------------------------------------------------------------------------------------------
    ' update application correspondence salutation
    '------------------------------------------------------------------------------------------
    Dim intCustomerRoleType As Integer, intCustomerOrder As Integer
    
    'APS 03/04/00
    Dim xmlNode As IXMLDOMNode
    Set xmlNode = vxmlRequest.selectSingleNode(".//CUSTOMERROLETYPE")
    If Not xmlNode Is Nothing And xmlNode.Text <> "" Then
        intCustomerRoleType = CInt(xmlNode.Text)
    Else
        intCustomerRoleType = 1
    End If
    Set xmlNode = vxmlRequest.selectSingleNode(".//CUSTOMERORDER")
    If Not xmlNode Is Nothing And xmlNode.Text <> "" Then
        intCustomerOrder = CInt(xmlNode.Text)
    Else
        intCustomerOrder = 1
    End If
    
    ' Update correspondence salutation only when RoleType = 1 and CustomerOrder <= 2
    If intCustomerRoleType = 1 And intCustomerOrder <= 2 Then
    'BMIDS00566 core AQR SYS4898 - DRC 20/06/02 - removed code to a seperate function
        UpdateAppCorresSalutation strApplicationNumber, intAFFNumber
    End If
    
    'EP2_56 GHun Update the RemovedToECustomer table when a customer is being added to a Transfer Of Equity case.
    strApplicationType = vxmlRequest.selectSingleNode(".//TYPEOFAPPLICATION").Text
    Set objComboDO = m_objContext.CreateInstance(gstrBASE_COMPONENT & ".ComboDO")
    If (objComboDO.IsItemInValidation("TypeOfMortgage", strApplicationType, "T") = True) Then
        
        'Build the Request node required for updating the RemovedToECustomer table
        Set xmlOut = xmlCreateDOMObject()
        Set xmlInputElem = m_objXmlAssist.GetRequestNode(vxmlRequest.ownerDocument)
        Set xmlTableElem = xmlOut.createElement("REMOVEDTOECUSTOMER")
        xmlInputElem.appendChild xmlTableElem
    
        strTag = "APPLICATIONNUMBER"
        Set xmlElem = xmlOut.createElement(strTag)
        xmlElem.Text = vxmlRequest.selectSingleNode(".//" & strTag).Text
        xmlTableElem.appendChild xmlElem
    
        strTag = "CIFNUMBER"
        Set xmlElem = xmlOut.createElement(strTag)
        xmlElem.Text = vxmlRequest.selectSingleNode(".//" & strTag).Text
        xmlTableElem.appendChild xmlElem
    
        Set xmlElem = xmlOut.createElement("OMIGACUSTOMERNUMBER")
        xmlElem.Text = vxmlRequest.selectSingleNode(".//CUSTOMERNUMBER").Text
        xmlTableElem.appendChild xmlElem
    
        strTag = "TITLE"
        Set xmlElem = xmlOut.createElement(strTag)
        xmlElem.Text = vxmlRequest.selectSingleNode(".//" & strTag).Text
        xmlTableElem.appendChild xmlElem
    
        strTag = "FIRSTFORENAME"
        Set xmlElem = xmlOut.createElement(strTag)
        xmlElem.Text = vxmlRequest.selectSingleNode(".//" & strTag).Text
        xmlTableElem.appendChild xmlElem
    
        strTag = "SECONDFORENAME"
        Set xmlElem = xmlOut.createElement(strTag)
        xmlElem.Text = vxmlRequest.selectSingleNode(".//" & strTag).Text
        xmlTableElem.appendChild xmlElem
    
        strTag = "OTHERFORENAMES"
        Set xmlElem = xmlOut.createElement(strTag)
        xmlElem.Text = vxmlRequest.selectSingleNode(".//" & strTag).Text
        xmlTableElem.appendChild xmlElem
    
        strTag = "SURNAME"
        Set xmlElem = xmlOut.createElement(strTag)
        xmlElem.Text = vxmlRequest.selectSingleNode(".//" & strTag).Text
        xmlTableElem.appendChild xmlElem
        
        Set xmlElem = xmlOut.createElement("CUSTOMERROLE")
        xmlElem.Text = vxmlRequest.selectSingleNode(".//CUSTOMERROLETYPE").Text
        xmlTableElem.appendChild xmlElem
    
        Set xmlElem = xmlOut.createElement("TYPE")
        xmlElem.Text = "A"
        xmlTableElem.appendChild xmlElem
    
        If m_objContext Is Nothing Then
            Set objIApplManagerDO = New ApplicationManagerDO
        Else
            Set objIApplManagerDO = m_objContext.CreateInstance(cstrCOMPONENT_NAME & ".ApplicationManagerDO")
        End If
        objIApplManagerDO.CreateRemovedToECustomer xmlInputElem
    End If
    'EP2_56 End
    
    If Not m_objContext Is Nothing Then
        m_objContext.SetComplete
    End If
    
    Set objIApplicationDO = Nothing
    Set objIADOAssist = Nothing
    Set xmlInputElem = Nothing
    Set xmlTableElem = Nothing
    Set xmlElem = Nothing
    Set xmlTempResponseNode = Nothing
    Set xmlResponseElem = Nothing
    'EP2_56 GHun
    Set objComboDO = Nothing
    Set xmlOut = Nothing
    Set objIApplManagerDO = Nothing
    'EP2_56 End
    
    Exit Sub
MaintainAddCustomerToApplicationVbErr:

    Set objIApplicationDO = Nothing
    Set objIADOAssist = Nothing
    Set xmlInputElem = Nothing
    Set xmlTableElem = Nothing
    Set xmlElem = Nothing
    Set xmlTempResponseNode = Nothing
    Set xmlResponseElem = Nothing
    'EP2_56 GHun
    Set objComboDO = Nothing
    Set xmlOut = Nothing
    Set objIApplManagerDO = Nothing
    'EP2_56 End
    
    Err.Raise Err.Number, Err.Source, Err.Description
    
End Sub

Private Sub MaintainDeleteCustomerFromApplication(ByVal vxmlRequest As IXMLDOMElement)
'----------------------------------------------------------------------------------
'Description : This function performs all the updations to maintain the customer
'              roles after adding customer to an application.
'              N.B. A seperate function is written because of the performance
'              problems.
'----------------------------------------------------------------------------------
On Error GoTo DeleteVbErr

    Const strFunctionName As String = "MaintainDeleteCustomerFromApplication"

    Dim xmlInputElem As IXMLDOMElement 'Input to ApplicationFactFind
    Dim xmlElem As IXMLDOMElement, xmlTableElem As IXMLDOMElement
    Dim xmlTempResponseNode As IXMLDOMElement, xmlResponseElem As IXMLDOMElement
    
    Dim strApplCount As String, strApplicationNumber As String
    Dim intAFFNumber As Integer
    
    Dim objIApplicationDO As IApplicationDO
    Dim objIADOAssist As IADOAssist
    Dim cmd As ADODB.Command, Cn As ADODB.Connection
    
    Set xmlResponseElem = vxmlRequest.ownerDocument.createElement("RESPONSE")
    xmlResponseElem.setAttribute "TYPE", "SUCCESS"
    
    '   Build Request node with either 'NumberOfApplicants' or 'NumberOfGuarantors'
    '   based on which ever is to be updated (not empty in the request)
    Set xmlInputElem = m_objXmlAssist.GetRequestNode(vxmlRequest.ownerDocument)
    Set xmlTableElem = vxmlRequest.ownerDocument.createElement("APPLICATIONFACTFIND")
    xmlInputElem.appendChild xmlTableElem
    
    Set xmlElem = vxmlRequest.ownerDocument.createElement("APPLICATIONNUMBER")
    strApplicationNumber = vxmlRequest.selectSingleNode(".//" & "APPLICATIONNUMBER").Text
    xmlElem.Text = strApplicationNumber
    xmlTableElem.appendChild xmlElem
    
    Set xmlElem = vxmlRequest.ownerDocument.createElement("APPLICATIONFACTFINDNUMBER")
    intAFFNumber = vxmlRequest.selectSingleNode(".//" & "APPLICATIONFACTFINDNUMBER").Text
    xmlElem.Text = intAFFNumber
    xmlTableElem.appendChild xmlElem
    
    strApplCount = m_objXmlAssist.GetTagValue(vxmlRequest, "NUMBEROFAPPLICANTS")
    If strApplCount <> "" Then
        Set xmlElem = vxmlRequest.ownerDocument.createElement("NUMBEROFAPPLICANTS")
        xmlElem.Text = strApplCount
        xmlTableElem.appendChild xmlElem
    Else
        strApplCount = m_objXmlAssist.GetTagValue(vxmlRequest, "NUMBEROFGUARANTORS")
        Set xmlElem = vxmlRequest.ownerDocument.createElement("NUMBEROFGUARANTORS")
        xmlElem.Text = strApplCount
        xmlTableElem.appendChild xmlElem
    End If
    
    If m_objContext Is Nothing Then
        Set objIApplicationDO = New ApplicationDO
        Set objIADOAssist = New ADOAssist
    Else
        Set objIApplicationDO = m_objContext.CreateInstance(App.Title & ".ApplicationDO")
        Set objIADOAssist = m_objContext.CreateInstance(gstrBASE_COMPONENT & ".ADOAssist")
    End If
    
    objIApplicationDO.UpdateApplicationFactFind xmlInputElem
     
    '------------------------------------------------------------------------------
    ' Validate the change
    '------------------------------------------------------------------------------
    ' Create the Command object, assign values to the parameters and execute it
    Set Cn = New ADODB.Connection
    Cn.ConnectionString = objIADOAssist.GetConnStr
    Cn.CursorLocation = adUseClient
    Cn.open

    Set cmd = New ADODB.Command
    With cmd
        Set .ActiveConnection = Cn
        .CommandType = adCmdStoredProc
        'ASt 15/03/01 CC012 Compatible with SQL Server and Oracle
#If GENERIC_SQL Then
            If objIADOAssist.GetDBEngine = Oracle Then
                .CommandText = "sp_CustomerReg.MaintainDelCustomerFromAppl"
            ElseIf objIADOAssist.GetDBEngine = SQLServer Then
                .CommandText = "usp_CustomerRegMaintainDelCustomerFromAppl"
            End If
#Else
            .CommandText = "sp_CustomerReg.MaintainDelCustomerFromAppl"
#End If
        .Parameters.Append .CreateParameter(, adBSTR, adParamInput, 12, strApplicationNumber)
        .Parameters.Append .CreateParameter(, adInteger, adParamInput, 5, intAFFNumber)
        .Execute
    End With
    '   Check for any errors raised
    Dim lngErrNo As Long
    If cmd.ActiveConnection.Errors.Count <> 0 Then
        lngErrNo = cmd.ActiveConnection.Errors.Item(0).Number
        m_objErrAssist.ThrowError TypeName(Me), strFunctionName, Err.Number
    End If

    Set cmd.ActiveConnection = Nothing
        Cn.Close
    Set Cn = Nothing

   
    If Not m_objContext Is Nothing Then
        m_objContext.SetComplete
    End If
    
    Set objIApplicationDO = Nothing
    Set objIADOAssist = Nothing
    Set cmd = Nothing
    Set Cn = Nothing
    Set objIADOAssist = Nothing
    Set xmlInputElem = Nothing
    Set xmlTableElem = Nothing
    Set xmlElem = Nothing
    Set xmlTempResponseNode = Nothing
    Set xmlResponseElem = Nothing
    
    Exit Sub
DeleteVbErr:
    ' Close the connection if still open
    If Not Cn Is Nothing Then
        If Cn.State = adStateOpen Then
                Cn.Close
        End If
    End If

    Set objIApplicationDO = Nothing
    Set objIADOAssist = Nothing
    Set cmd = Nothing
    Set Cn = Nothing
    Set xmlInputElem = Nothing
    Set xmlTableElem = Nothing
    Set xmlElem = Nothing
    Set xmlTempResponseNode = Nothing
    Set xmlResponseElem = Nothing
    
    Err.Raise Err.Number, Err.Source, Err.Description
    
End Sub

Private Function IApplicationManagerTxBO_CreateCustomerLockApplicationLock(ByVal vxmlRequest As IXMLDOMElement) As IXMLDOMNode
' header ----------------------------------------------------------------------------------
' description:  create single instance of the persistant data associated with thie table
'               'CustomerLockApplicationLock'
'               invokes appropriate Data Object(s) to perform persistance
' pass:         vxmlRequest  xml Request element
' return:       xml Response Node
'------------------------------------------------------------------------------------------
On Error GoTo CreateVbErr

    Const strFunctionName As String = "IApplicationManagerTxBO_CreateCustomerLockApplicationLock"

    Dim xmlElement As IXMLDOMElement
    Dim xmlOut As New FreeThreadedDOMDocument40
    xmlOut.validateOnParse = False
    xmlOut.setProperty "NewParser", True
    Dim xmlResponseElem As IXMLDOMElement

    Dim objIApplicationManagerDO As IApplicationManagerDO
    
    Set xmlResponseElem = xmlOut.createElement("RESPONSE")
    xmlOut.appendChild xmlResponseElem
    xmlResponseElem.setAttribute "TYPE", "SUCCESS"
                
    ' Check if the element passed in is the one with the table name else
    ' look below the node that is passed in
    If vxmlRequest.nodeName = cstrTABLE_CUSTLOCKAPPLLOCK Then
        Set xmlElement = vxmlRequest
    Else
        Set xmlElement = vxmlRequest.getElementsByTagName(cstrTABLE_CUSTLOCKAPPLLOCK).Item(0)
    End If
    
    If xmlElement Is Nothing Then
        m_objErrAssist.ThrowError TypeName(Me), _
                                  strFunctionName, _
                                  oeMissingPrimaryTag, _
                                  cstrTABLE_CUSTLOCKAPPLLOCK & " tag not found"
    End If
   
    If m_objContext Is Nothing Then
        Set objIApplicationManagerDO = New ApplicationManagerDO
    Else
        Set objIApplicationManagerDO = _
                    m_objContext.CreateInstance(App.Title & ".ApplicationManagerDO")
    End If
    
    objIApplicationManagerDO.CreateCustomerLockApplicationLock xmlElement
    
    Set IApplicationManagerTxBO_CreateCustomerLockApplicationLock = xmlResponseElem
    
    If Not m_objContext Is Nothing Then
        m_objContext.SetComplete
    End If
    
CreateExit:

    Set objIApplicationManagerDO = Nothing
    Set xmlElement = Nothing
    Set xmlOut = Nothing
    Set xmlResponseElem = Nothing
    
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

    Set IApplicationManagerTxBO_CreateCustomerLockApplicationLock = _
                                        m_objErrAssist.CreateErrorResponseEx()

    ' If we get duplicate key do not abort the transaction but allow objects higher
    ' up the chain to abort their transactions if they want to
    If Err.Number = omiga4DuplicateKey Then
        If Not m_objContext Is Nothing Then
            m_objContext.SetComplete
        End If
    Else
        If Not m_objContext Is Nothing Then
            m_objContext.SetAbort
        End If
    End If
    
    ' go to clean-up section
    Resume CreateExit

End Function

Private Function IApplicationManagerTxBO_CreateCustomerRole( _
                ByVal vxmlRequest As IXMLDOMElement) As IXMLDOMNode
' header ----------------------------------------------------------------------------------
' description:  create multiple instances of the persistant data associated with this
'               business object
'               invokes appropriate Data Object(s) to perform persistance
' pass:         vxmlRequest  xml Request element
' return:       xml Response Node
'------------------------------------------------------------------------------------------
On Error GoTo CreateVbErr

    Const strFunctionName As String = "IApplicationManagerTxBO_CreateCustomerRole"

    Dim xmlNodeList As IXMLDOMNodeList
    Dim xmlElement As IXMLDOMElement
    Dim xmlOut As New FreeThreadedDOMDocument40
    xmlOut.validateOnParse = False
    xmlOut.setProperty "NewParser", True
    Dim xmlResponseElem As IXMLDOMElement
   
    Dim objIApplManagerDO As IApplicationManagerDO
    
    Set xmlResponseElem = xmlOut.createElement("RESPONSE")
    xmlOut.appendChild xmlResponseElem
    xmlResponseElem.setAttribute "TYPE", "SUCCESS"
                
    ' Check if the element passed in is the one with the table name else
    ' look below the node that is passed in
    If vxmlRequest.nodeName = cstTABLE_CUSTOMERROLE Then
        Set xmlElement = vxmlRequest
    Else
        Set xmlNodeList = vxmlRequest.getElementsByTagName(cstTABLE_CUSTOMERROLE)
    End If
    
    If xmlElement Is Nothing Then
        If xmlNodeList.length = 0 Then
            m_objErrAssist.ThrowError TypeName(Me), _
                                  strFunctionName, _
                                  oeMissingPrimaryTag, _
                                  cstTABLE_CUSTOMERROLE & " tag not found"
        End If
    End If
   
    If m_objContext Is Nothing Then
        Set objIApplManagerDO = New ApplicationManagerDO
    Else
        Set objIApplManagerDO = m_objContext.CreateInstance( _
                            App.Title & ".ApplicationManagerDO")
    End If
    
    'JLD check input xml to decide whether to create locks or not
    Dim bCreateLock As Boolean
    Dim strLocking As String
    strLocking = m_objXmlAssist.GetTagValue(vxmlRequest, "LOCKING", , True)
    If strLocking = "1" Then
        bCreateLock = True
    Else
        bCreateLock = False
    End If
    If Not xmlElement Is Nothing Then
        objIApplManagerDO.CreateCustomerRole xmlElement
        If bCreateLock = True Then
            CreateCustLockApplLock xmlElement
        End If
    Else
        For Each xmlElement In xmlNodeList
            objIApplManagerDO.CreateCustomerRole xmlElement
            If bCreateLock = True Then
                CreateCustLockApplLock xmlElement
            End If
        Next xmlElement
    End If
    
    Set IApplicationManagerTxBO_CreateCustomerRole = xmlResponseElem
    
    If Not m_objContext Is Nothing Then
        m_objContext.SetComplete
    End If
    
CreateExit:

    Set objIApplManagerDO = Nothing
    Set xmlNodeList = Nothing
    Set xmlElement = Nothing
    Set xmlOut = Nothing
    Set xmlResponseElem = Nothing
    
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

    Set IApplicationManagerTxBO_CreateCustomerRole = m_objErrAssist.CreateErrorResponseEx()

    If Not m_objContext Is Nothing Then
        m_objContext.SetAbort
    End If
    
    ' go to clean-up section
    Resume CreateExit
End Function

Private Function IApplicationManagerTxBO_CreateFullApplicationDetails( _
                ByVal vxmlRequest As IXMLDOMElement) As IXMLDOMNode
' header -------------------------------------------------------------------------
' description:
' pass:         vxmlRequest  xml Request element
' return:       xml Response Node
'---------------------------------------------------------------------------------
On Error GoTo CreateVbErr

    Const strFunctionName As String = "IApplicationManagerTxBO_CreateFullApplicationDetails"

    Dim xmlElement As IXMLDOMElement
    Dim xmlOut As New FreeThreadedDOMDocument40
    xmlOut.validateOnParse = False
    xmlOut.setProperty "NewParser", True
    Dim xmlResponseElem As IXMLDOMElement
    Dim xmlTempResponseNode As IXMLDOMElement
    
    Dim objIApplicationBO As IApplicationBO
    Dim objINewPropertyBO As INewPropertyBO
    
    Set xmlResponseElem = xmlOut.createElement("RESPONSE")
    xmlOut.appendChild xmlResponseElem
    xmlResponseElem.setAttribute "TYPE", "SUCCESS"
                
    If vxmlRequest.nodeName = "APPLICATION" Then
        Set xmlElement = vxmlRequest
    Else
        Set xmlElement = vxmlRequest.getElementsByTagName("APPLICATION").Item(0)
    End If
    
    If xmlElement Is Nothing Then
        m_objErrAssist.ThrowError TypeName(Me), strFunctionName, _
                            oeMissingPrimaryTag, "APPLICATION" & " tag not found"
    End If
   
    If m_objContext Is Nothing Then
        Set objIApplicationBO = New ApplicationBO
        Set objINewPropertyBO = New NewPropertyBO
    Else
        Set objIApplicationBO = _
                m_objContext.CreateInstance(cstrCOMPONENT_NAME & ".ApplicationBO")
        Set objINewPropertyBO = _
                m_objContext.CreateInstance(cstrCOMPONENT_NAME & ".NewPropertyBO")
    End If
    
    Set xmlTempResponseNode = objIApplicationBO.CreateApplication(vxmlRequest)
    m_objErrAssist.CheckResponse xmlTempResponseNode, xmlResponseElem, True
    
    Dim strApplNumber As String, strAFFNumber As String
    strApplNumber = xmlTempResponseNode.selectSingleNode(".//APPLICATIONNUMBER").Text
    strAFFNumber = xmlTempResponseNode.selectSingleNode(".//APPLICATIONFACTFINDNUMBER").Text
        
    'Update ApplicationFactFind element with generated keys
    Call SetInputForCreateFullApplication(vxmlRequest, "APPLICATIONFACTFIND", _
                                        strApplNumber, strAFFNumber)
        
     ' Create ApplicationVerification
     Call SetInputForCreateFullApplication(vxmlRequest, "APPLICATIONVERIFICATION", _
                                        strApplNumber, strAFFNumber)

    Set xmlTempResponseNode = objIApplicationBO.CreateApplicationVerification(vxmlRequest)
    m_objErrAssist.CheckResponse xmlTempResponseNode, xmlResponseElem, True
        
    ' Creare FinancialSummary
    Call SetInputForCreateFullApplication(vxmlRequest, "FINANCIALSUMMARY", _
                                          strApplNumber, strAFFNumber)
    Set xmlTempResponseNode = objIApplicationBO.CreateFinancialSummary(vxmlRequest)
    m_objErrAssist.CheckResponse xmlTempResponseNode, xmlResponseElem, True
    
    ' MC 20/06/2000 Create NewLoan
    Set xmlElement = vxmlRequest.selectSingleNode(".//NEWLOAN")
    If Not xmlElement Is Nothing Then
        Call SetInputForCreateFullApplication(vxmlRequest, "NEWLOAN", _
                                              strApplNumber, strAFFNumber)
        Set xmlTempResponseNode = objIApplicationBO.SaveNewLoan(vxmlRequest)
        m_objErrAssist.CheckResponse xmlTempResponseNode, xmlResponseElem, True
    End If
    
    ' Create NewPropertyDetails
    Call SetInputForCreateFullApplication(vxmlRequest, "NEWPROPERTY", _
                                          strApplNumber, strAFFNumber)
    Call SetInputForCreateFullApplication(vxmlRequest, "NEWPROPERTYDEPOSIT", _
                                          strApplNumber, strAFFNumber, False)
    Call SetInputForCreateFullApplication(vxmlRequest, "NEWPROPERTYLEASEHOLD", _
                                          strApplNumber, strAFFNumber, False)
    Call SetInputForCreateFullApplication(vxmlRequest, "NEWPROPERTYROOMTYPE", _
                                          strApplNumber, strAFFNumber, False)
    Call SetInputForCreateFullApplication(vxmlRequest, "SHAREDOWNERSHIPDETAILS", _
                                          strApplNumber, strAFFNumber, False)
    Set xmlTempResponseNode = objINewPropertyBO.CreateNewPropertyDetails(vxmlRequest)
    m_objErrAssist.CheckResponse xmlTempResponseNode, xmlResponseElem, True
    
    ' Create NewPropertyAddress
    Call SetInputForCreateFullApplication(vxmlRequest, "NEWPROPERTYADDRESS", _
                                          strApplNumber, strAFFNumber)
    Set xmlTempResponseNode = objINewPropertyBO.CreateNewPropertyAddress(vxmlRequest)
    m_objErrAssist.CheckResponse xmlTempResponseNode, xmlResponseElem, True
    
    ' Add ApplicationNumber and ApplicationFactFindNumber to Response
    Set xmlElement = xmlOut.createElement("APPLICATIONNUMBER")
    xmlElement.Text = strApplNumber
    xmlResponseElem.appendChild xmlElement
    
    Set xmlElement = xmlOut.createElement("APPLICATIONFACTFINDNUMBER")
    xmlElement.Text = strAFFNumber
    xmlResponseElem.appendChild xmlElement
    
    Set IApplicationManagerTxBO_CreateFullApplicationDetails = xmlResponseElem
    
    If Not m_objContext Is Nothing Then
        m_objContext.SetComplete
    End If
    
CreateExit:
    
    Set objIApplicationBO = Nothing
    Set objINewPropertyBO = Nothing
    Set xmlElement = Nothing
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

    Set IApplicationManagerTxBO_CreateFullApplicationDetails = m_objErrAssist.CreateErrorResponseEx()

    If Not m_objContext Is Nothing Then
        m_objContext.SetAbort
    End If
    
    ' go to clean-up section
    Resume CreateExit

End Function

Public Function IApplicationManagerTxBO_CreateLock( _
                        ByVal vxmlRequest As IXMLDOMElement) As IXMLDOMNode
' header ----------------------------------------------------------------------------------
' description:
'   Create a single application lock.
' pass:
'   vstrXMLRequest
'       xml Request data stream containing data to be persisted
'       Format:
'            <REQUEST>
'                <APPLICATIONLOCK>
'                    <APPLICATIONNUMBER></APPLICATIONNUMBER>
'                    <UNITID></UNITID>
'                    <USERID></USERID>
'                    <LOCKDATE></LOCKDATE>
'                    <TYPEOFLOCK></TYPEOFLOCK>
'                    <MACHINEID></MACHINEID>
'                    <CHANNELID></CHANNELID>
'                </APPLICATIONLOCK>
'            </REQUEST>
' return: n/a
'------------------------------------------------------------------------------------------
' RF 02/12/99 AQR SC017: Enhancement - validate working hours on creating an application lock.
'                        Function now only copes with creating a single lock for simplicity.
'------------------------------------------------------------------------------------------
On Error GoTo CreateLockVbErr

    Const strFunctionName As String = "IApplicationManagerTxBO_CreateLock"

    Dim xmlElement As IXMLDOMElement
    Dim xmlOut As New FreeThreadedDOMDocument40
    xmlOut.validateOnParse = False
    xmlOut.setProperty "NewParser", True
    Dim xmlRequestElem As IXMLDOMElement
    Dim xmlResponseElem As IXMLDOMElement
    Dim xmlAccessAuditTableNode As IXMLDOMNode
    Dim xmlTempElem As IXMLDOMElement
        
    Dim objIApplicationManagerDO As IApplicationManagerDO
    Dim objIAuditDO As IAuditDO
    Dim objAuditBO As AuditBO
    
    Set xmlResponseElem = xmlOut.createElement("RESPONSE")
    xmlOut.appendChild xmlResponseElem
    xmlResponseElem.setAttribute "TYPE", "SUCCESS"
                
    ' Check if the element passed in is the one with the table name else
    ' look below the node that is passed in
    If vxmlRequest.nodeName = cstrTABLE_APPLLOCK Then
        Set xmlElement = vxmlRequest
    Else
        Set xmlElement = vxmlRequest.getElementsByTagName(cstrTABLE_APPLLOCK).Item(0)
    End If
    
    If xmlElement Is Nothing Then
        m_objErrAssist.ThrowError TypeName(Me), _
                                  strFunctionName, _
                                  oeMissingPrimaryTag, _
                                  cstrTABLE_APPLLOCK & " tag not found"
    End If
   
    If m_objContext Is Nothing Then
        Set objIApplicationManagerDO = New ApplicationManagerDO
        Set objIAuditDO = New AuditDO
        Set objAuditBO = New AuditBO

    Else
        Set objIApplicationManagerDO = m_objContext.CreateInstance(App.Title & ".ApplicationManagerDO")
        Set objIAuditDO = m_objContext.CreateInstance(gstrAUDIT_COMPONENT & ".AuditDO")
        Set objAuditBO = m_objContext.CreateInstance(gstrAUDIT_COMPONENT & ".AuditBO")
    End If
    
    Dim strResponse As String
    Dim blnSuccessInd As Boolean
    Dim lngValidateWorkingHoursErr As Long
    
    '------------------------------------------------------------------------------------------
    ' RF 02/12/99 AQR SC017: Enhancement - validate working hours
    '------------------------------------------------------------------------------------------
    Err.Clear
    On Error Resume Next
    m_objXmlAssist.ChangeNodeName vxmlRequest, "APPLICATIONLOCK", "OMIGAUSER"
    ValidateWorkingHours vxmlRequest.xml
    m_objErrAssist.SaveErr
    On Error GoTo CreateLockVbErr
    
    lngValidateWorkingHoursErr = m_objErrAssist.LastErr
    
    If lngValidateWorkingHoursErr = 0 Then
        blnSuccessInd = True
    Else
        If m_objErrAssist.IsWarning(lngValidateWorkingHoursErr) = True Then
            blnSuccessInd = True
            m_objErrAssist.ReRaise
        Else
            blnSuccessInd = False
        End If
    End If
    
    If blnSuccessInd = True Then
        '------------------------------------------------------------------------------------------
        ' do the create
        '------------------------------------------------------------------------------------------
        Err.Clear
        On Error Resume Next
        objIApplicationManagerDO.CreateLock xmlElement
        m_objErrAssist.SaveErr
        On Error GoTo CreateLockVbErr
        
        ' Check if already locked. If so raise error if not locked to this user
        If m_objErrAssist.LastErr <> 0 Then
            If m_objErrAssist.LastErr = omiga4DuplicateKey Then
            
                ' APS UNIT TEST REF 102 10/09/99
                'If IsLockedToDifferentUser(xmlElem) = True Then
                    
                    '****************************************************************
                    ' FIXME - Need to be able to substitute Application number into
                    '         message 182
                    '****************************************************************
                    
                    m_objErrAssist.ThrowError TypeName(Me), strFunctionName, 182
                'End If
                
            Else
                m_objErrAssist.ReRaise
            End If
        End If
    End If
    
        '------------------------------------------------------------------------------------------
    ' RF 29/09/99
    ' On creation of an application lock, write an access audit record
    '------------------------------------------------------------------------------------------
    
    Set xmlRequestElem = m_objXmlAssist.GetRequestNode(vxmlRequest.ownerDocument)
    m_objXmlAssist.CopyNode "OMIGAUSER", "ACCESSAUDIT", vxmlRequest, xmlRequestElem, True
'    m_objXmlAssist.ChangeNodeName vxmlRequest, "OMIGAUSER", "ACCESSAUDIT"
     
    Set xmlAccessAuditTableNode = _
        xmlRequestElem.getElementsByTagName("ACCESSAUDIT").Item(0)
    
    Set xmlTempElem = xmlOut.createElement("AUDITRECORDTYPE")
    xmlTempElem.Text = objIAuditDO.GetApplicationLockValueId()
    xmlAccessAuditTableNode.appendChild xmlTempElem
    
    Set xmlTempElem = xmlOut.createElement("SUCCESSINDICATOR")
    If blnSuccessInd = True Then
        xmlTempElem.Text = "1"
    Else
        xmlTempElem.Text = "0"
    End If
    xmlAccessAuditTableNode.appendChild xmlTempElem
    
    strResponse = objAuditBO.CreateAccessAudit(xmlRequestElem.xml)
    
    m_objErrAssist.CheckXMLResponse strResponse, True
    
    '------------------------------------------------------------------------------------------
    ' ensure any error from ValidateWorkingHours gets raised
    '------------------------------------------------------------------------------------------
    
    If blnSuccessInd = False Then
        m_objErrAssist.RaiseError TypeName(Me), strFunctionName, lngValidateWorkingHoursErr
    End If
    
    Set IApplicationManagerTxBO_CreateLock = xmlResponseElem
    
    If Not m_objContext Is Nothing Then
        m_objContext.SetComplete
    End If
    
CreateLockExit:

    Set objIApplicationManagerDO = Nothing
    Set xmlElement = Nothing
    Set xmlOut = Nothing
    Set xmlResponseElem = Nothing
    Set xmlRequestElem = Nothing
    Set xmlAccessAuditTableNode = Nothing
    Set xmlTempElem = Nothing

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

    Set IApplicationManagerTxBO_CreateLock = m_objErrAssist.CreateErrorResponseEx()

    If Not m_objContext Is Nothing Then
        m_objContext.SetAbort
    End If
    
    ' go to clean-up section
    Resume CreateLockExit
End Function

Public Sub ValidateWorkingHours(ByVal vstrXMLRequest As String)
' header ----------------------------------------------------------------------------------
' description:
' pass:
'   vstrXMLRequest
'   Format:
'       <REQUEST>
'           <OMIGAUSER>
'               <USERID></USERID>
'               <CHANNELID></CHANNELID>
'           </OMIGAUSER>
'       </REQUEST>
' return: n/a
'------------------------------------------------------------------------------------------
On Error GoTo ValidateWorkingHoursVbErr

    Const strFunctionName As String = "ValidateWorkingHours"
    
    Dim objOrgBO As OrganisationBO
 
    If m_objContext Is Nothing Then
        Set objOrgBO = New OrganisationBO
    Else
        Set objOrgBO = m_objContext.CreateInstance(gstrORGANISATION_COMPONENT & ".OrganisationBO")
    End If
    
    objOrgBO.ValidateWorkingHours vstrXMLRequest

    Set objOrgBO = Nothing
    
    Exit Sub

ValidateWorkingHoursVbErr:
   
    Set objOrgBO = Nothing
   
    m_objErrAssist.AddToErrSource strFunctionName
        
    If m_objErrAssist.IsSystemError = True Then
        m_objErrAssist.LogError TypeName(Me), strFunctionName, Err.Number, Err.Description
    End If
        
    If Not m_objContext Is Nothing Then
        m_objContext.SetAbort
    End If
    
    Err.Raise Err.Number, Err.Source, Err.Description

End Sub

Private Function IApplicationManagerTxBO_DeleteAllCustomerLockApplicationLock(ByVal vxmlRequest As IXMLDOMElement) As IXMLDOMNode
' header ----------------------------------------------------------------------------------
' description:  Delete multiple instances of the persistant data associated with the table
'               'ApplicationLock'
'               invokes appropriate Data Object(s) to perform persistance
' pass:         vxmlRequest  xml Request element
' return:       xml Response Node
'------------------------------------------------------------------------------------------
On Error GoTo DeleteVbErr

    Const strFunctionName As String = "IApplicationManagerTxBO_DeleteAllCustomerLockApplicationLock"

    Dim xmlElement As IXMLDOMElement
    Dim xmlOut As New FreeThreadedDOMDocument40
    xmlOut.validateOnParse = False
    xmlOut.setProperty "NewParser", True
    Dim xmlResponseElem As IXMLDOMElement
    
    Dim objIApplManagerDO As IApplicationManagerDO
    
    Set xmlResponseElem = xmlOut.createElement("RESPONSE")
    xmlOut.appendChild xmlResponseElem
    xmlResponseElem.setAttribute "TYPE", "SUCCESS"
                
    ' Check if the element passed in is the one with the table name else
    ' look below the node that is passed in
    If vxmlRequest.nodeName = cstrTABLE_CUSTLOCKAPPLLOCK Then
        Set xmlElement = vxmlRequest
    Else
        Set xmlElement = vxmlRequest.getElementsByTagName(cstrTABLE_CUSTLOCKAPPLLOCK).Item(0)
    End If
    
    If xmlElement Is Nothing Then
        m_objErrAssist.ThrowError TypeName(Me), _
                                  strFunctionName, _
                                  oeMissingPrimaryTag, _
                                  cstrTABLE_CUSTLOCKAPPLLOCK & " tag not found"
    End If
   
    If m_objContext Is Nothing Then
        Set objIApplManagerDO = New ApplicationManagerDO
    Else
        Set objIApplManagerDO = m_objContext.CreateInstance(App.Title & ".ApplicationManagerDO")
    End If
    
    objIApplManagerDO.DeleteAllCustomerLockApplicationLock xmlElement
    
    Set IApplicationManagerTxBO_DeleteAllCustomerLockApplicationLock = xmlResponseElem
    
    If Not m_objContext Is Nothing Then
        m_objContext.SetComplete
    End If
    
DeleteExit:

    Set objIApplManagerDO = Nothing
    Set xmlElement = Nothing
    Set xmlOut = Nothing
    Set xmlResponseElem = Nothing
    
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

    Set IApplicationManagerTxBO_DeleteAllCustomerLockApplicationLock = _
                                        m_objErrAssist.CreateErrorResponseEx()

    If Not m_objContext Is Nothing Then
        m_objContext.SetAbort
    End If
    
    ' go to clean-up section
    Resume DeleteExit
End Function

Private Function IApplicationManagerTxBO_DeleteCustomerFromApplication( _
                    ByVal vxmlRequest As IXMLDOMElement) As IXMLDOMNode
' header ----------------------------------------------------------------------------------
' description:
'   Remove the link between a customer and the application
' pass:
'   vXMLRequest
' return:   n/a
' Raise Errors:
'------------------------------------------------------------------------------------------
On Error GoTo DeleteCustomerFromApplicationVbErr

    Const strFunctionName As String = "IApplicationManagerTxBO_DeleteCustomerFromApplication"
    
    Dim xmlOut As FreeThreadedDOMDocument40
    Dim xmlResponseElem As IXMLDOMElement
    Dim xmlTempResponseNode As IXMLDOMNode
    Dim xmlInputElem As IXMLDOMElement
    Dim xmlCustRoleElem As IXMLDOMElement
    Dim xmlElement As IXMLDOMElement

    Dim objIApplManagerDO As IApplicationManagerDO
    Dim strApplicationType As String                 'BMIDS758
    Dim objComboDO As ComboDO                        'BMIDS758
    Dim xmlToEElem As IXMLDOMElement                 'BMIDS758
    
    Set xmlOut = New FreeThreadedDOMDocument40
    xmlOut.validateOnParse = False
    xmlOut.setProperty "NewParser", True
    Set xmlResponseElem = xmlOut.createElement("RESPONSE")
    xmlOut.appendChild xmlResponseElem
    xmlResponseElem.setAttribute "TYPE", "SUCCESS"
    
    If m_objContext Is Nothing Then
        Set objIApplManagerDO = New ApplicationManagerDO
    Else
        Set objIApplManagerDO = _
            m_objContext.CreateInstance(cstrCOMPONENT_NAME & ".ApplicationManagerDO")
    End If
    
    Set objComboDO = m_objContext.CreateInstance(gstrBASE_COMPONENT & ".ComboDO")   ' BMIDS758
    
    '--------------------------------------------
    ' delete CustomerRole
    '---------------------------------------------
    '   Build the Request node required for deletion of CustomerRole
    
    Set xmlInputElem = m_objXmlAssist.GetRequestNode(vxmlRequest.ownerDocument)
    Set xmlCustRoleElem = xmlOut.createElement("CUSTOMERROLE")
    xmlInputElem.appendChild xmlCustRoleElem
    
    Dim strTag As String
    strTag = "APPLICATIONNUMBER"
    Set xmlElement = xmlOut.createElement(strTag)
    xmlElement.Text = vxmlRequest.selectSingleNode(".//" & strTag).Text
    xmlCustRoleElem.appendChild xmlElement
    
    strTag = "APPLICATIONFACTFINDNUMBER"
    Set xmlElement = xmlOut.createElement(strTag)
    xmlElement.Text = vxmlRequest.selectSingleNode(".//" & strTag).Text
    xmlCustRoleElem.appendChild xmlElement
    
    strTag = "CUSTOMERNUMBER"
    Set xmlElement = xmlOut.createElement(strTag)
    xmlElement.Text = vxmlRequest.selectSingleNode(".//" & strTag).Text
    xmlCustRoleElem.appendChild xmlElement
    
    strTag = "CUSTOMERVERSIONNUMBER"
    Set xmlElement = xmlOut.createElement(strTag)
    xmlElement.Text = vxmlRequest.selectSingleNode(".//" & strTag).Text
    xmlCustRoleElem.appendChild xmlElement
    
    objIApplManagerDO.DeleteCustomerRole xmlInputElem
    
    ' Delete the locks (records from link tables), with respect to this customer
    DeleteCustomerLocks xmlInputElem
    
    'BMIDS758
    'Update the RemovedToECustomer table when a customer is being deleted from a Transfer Of Equity case.
    
    strApplicationType = vxmlRequest.selectSingleNode(".//TYPEOFAPPLICATION").Text
    If (objComboDO.IsItemInValidation("TypeOfMortgage", strApplicationType, "CC") = True) Then   'BMIDS836
        
        'Build the Request node required for updating the RemovedToECustomer table
    
        Set xmlInputElem = m_objXmlAssist.GetRequestNode(vxmlRequest.ownerDocument)
        Set xmlToEElem = xmlOut.createElement("REMOVEDTOECUSTOMER")
        xmlInputElem.appendChild xmlToEElem
    
        strTag = "APPLICATIONNUMBER"
        Set xmlElement = xmlOut.createElement(strTag)
        xmlElement.Text = vxmlRequest.selectSingleNode(".//" & strTag).Text
        xmlToEElem.appendChild xmlElement
    
        strTag = "CIFNUMBER"
        Set xmlElement = xmlOut.createElement(strTag)
        xmlElement.Text = vxmlRequest.selectSingleNode(".//" & strTag).Text
        xmlToEElem.appendChild xmlElement
    
        strTag = "OMIGACUSTOMERNUMBER"
        Set xmlElement = xmlOut.createElement(strTag)
        xmlElement.Text = vxmlRequest.selectSingleNode(".//" & strTag).Text
        xmlToEElem.appendChild xmlElement
    
        strTag = "TITLE"
        Set xmlElement = xmlOut.createElement(strTag)
        xmlElement.Text = vxmlRequest.selectSingleNode(".//" & strTag).Text
        xmlToEElem.appendChild xmlElement
    
        strTag = "FIRSTFORENAME"
        Set xmlElement = xmlOut.createElement(strTag)
        xmlElement.Text = vxmlRequest.selectSingleNode(".//" & strTag).Text
        xmlToEElem.appendChild xmlElement
    
        strTag = "SECONDFORENAME"
        Set xmlElement = xmlOut.createElement(strTag)
        xmlElement.Text = vxmlRequest.selectSingleNode(".//" & strTag).Text
        xmlToEElem.appendChild xmlElement
    
        strTag = "OTHERFORENAMES"
        Set xmlElement = xmlOut.createElement(strTag)
        xmlElement.Text = vxmlRequest.selectSingleNode(".//" & strTag).Text
        xmlToEElem.appendChild xmlElement
    
        strTag = "SURNAME"
        Set xmlElement = xmlOut.createElement(strTag)
        xmlElement.Text = vxmlRequest.selectSingleNode(".//" & strTag).Text
        xmlToEElem.appendChild xmlElement
        
        'EP2_56 GHun
        Set xmlElement = xmlOut.createElement("CUSTOMERROLE")
        xmlElement.Text = vxmlRequest.selectSingleNode(".//CUSTOMERROLETYPE").Text
        xmlToEElem.appendChild xmlElement
    
        Set xmlElement = xmlOut.createElement("TYPE")
        xmlElement.Text = "R"
        xmlToEElem.appendChild xmlElement
        'EP2_56 End
    
        objIApplManagerDO.CreateRemovedToECustomer xmlInputElem
    End If
        
    'BMIDS758  End
    
    '------------------------------------------------------------------
    ' validate the change
    '------------------------------------------------------------------
    MaintainDeleteCustomerFromApplication vxmlRequest


    '** COMMENTED FOR PERFORMANCE TEST
'    Set xmlTempResponseNode = _
'            IApplicationManagerTxBO_MaintainCustomerRoles(vxmlRequest)
'    m_objErrAssist.CheckResponse xmlTempResponseNode, xmlResponseElem, True
    
    Set IApplicationManagerTxBO_DeleteCustomerFromApplication = xmlResponseElem
    
    If Not m_objContext Is Nothing Then
        m_objContext.SetComplete
    End If
    
DeleteCustomerFromApplicationExit:
    Set xmlOut = Nothing
    Set xmlResponseElem = Nothing
    Set xmlTempResponseNode = Nothing
    Set xmlInputElem = Nothing
    Set xmlCustRoleElem = Nothing
    Set xmlElement = Nothing
    Set objIApplManagerDO = Nothing
    
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

    Set IApplicationManagerTxBO_DeleteCustomerFromApplication = _
                                m_objErrAssist.CreateErrorResponseEx()

    If Not m_objContext Is Nothing Then
        m_objContext.SetAbort
    End If
    
    ' go to clean-up section
    Resume DeleteCustomerFromApplicationExit
    
End Function

Private Function IApplicationManagerTxBO_DeleteCustomerLockApplicationLock(ByVal vxmlRequest As IXMLDOMElement) As IXMLDOMNode
' header ----------------------------------------------------------------------------------
' description:  Delete Single instances of the persistant data associated with the table
'               'CustomerLockApplicationLock'
'               invokes appropriate Data Object(s) to perform persistance
' pass:         vxmlRequest  xml Request element
' return:       xml Response Node
'------------------------------------------------------------------------------------------
On Error GoTo DeleteVbErr

    Const strFunctionName As String = "IApplicationManagerTxBO_DeleteCustomerLockApplicationLock"

    Dim xmlElement As IXMLDOMElement
    Dim xmlOut As New FreeThreadedDOMDocument40
    xmlOut.validateOnParse = False
    xmlOut.setProperty "NewParser", True
    Dim xmlResponseElem As IXMLDOMElement
    
    Dim objIApplicationManagerDO As IApplicationManagerDO
    
    Set xmlResponseElem = xmlOut.createElement("RESPONSE")
    xmlOut.appendChild xmlResponseElem
    xmlResponseElem.setAttribute "TYPE", "SUCCESS"
                
    ' Check if the element passed in is the one with the table name else
    ' look below the node that is passed in
    If vxmlRequest.nodeName = cstrTABLE_CUSTLOCKAPPLLOCK Then
        Set xmlElement = vxmlRequest
    Else
        Set xmlElement = vxmlRequest.getElementsByTagName(cstrTABLE_CUSTLOCKAPPLLOCK).Item(0)
    End If
    
    If xmlElement Is Nothing Then
        m_objErrAssist.ThrowError TypeName(Me), _
                                  strFunctionName, _
                                  oeMissingPrimaryTag, _
                                  cstrTABLE_CUSTLOCKAPPLLOCK & " tag not found"
    End If
   
    If m_objContext Is Nothing Then
        Set objIApplicationManagerDO = New ApplicationManagerDO
    Else
        Set objIApplicationManagerDO = m_objContext.CreateInstance(App.Title & ".ApplicationManagerDO")
    End If
    
    objIApplicationManagerDO.DeleteCustomerLockApplicationLock xmlElement
    
    Set IApplicationManagerTxBO_DeleteCustomerLockApplicationLock = xmlResponseElem
                        
DeleteExit:

    Set objIApplicationManagerDO = Nothing
    Set xmlElement = Nothing
    Set xmlOut = Nothing
    Set xmlResponseElem = Nothing
        
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

    Set IApplicationManagerTxBO_DeleteCustomerLockApplicationLock = _
                                    m_objErrAssist.CreateErrorResponseEx()

    If Not m_objContext Is Nothing Then
        m_objContext.SetAbort
    End If
    
    ' go to clean-up section
    Resume DeleteExit

End Function

Private Sub DeleteCustomerLocks(ByVal objCustomerRoleElem As IXMLDOMElement)
' header ----------------------------------------------------------------------------------
' description:  Deletes the locks held on the customer
' pass:         objCustomerRoleElem  Customer role element
'
' return:       n/a
'
' Raise Errors:
'------------------------------------------------------------------------------------------
On Error GoTo DeleteCustomerLocksVbErr

    Const strFunctionName As String = "DeleteCustomerLocks"

    Dim objRequestDoc As New FreeThreadedDOMDocument40
    objRequestDoc.validateOnParse = False
    objRequestDoc.setProperty "NewParser", True
    Dim objTableElem As IXMLDOMElement
    Dim objRequestNode As IXMLDOMNode
    Dim objElement As IXMLDOMElement
    Dim objIApplManagerDO As IApplicationManagerDO
    Dim objCustomerBO As CustomerBO

    If objCustomerRoleElem Is Nothing Then
        m_objErrAssist.RaiseError TypeName(Me), strFunctionName, oeInvalidParameter, _
                                "Customer Role element is set to nothing"
    End If

    Dim strApplicationNo As String
    Dim strCustomerNo As String

    strApplicationNo = m_objXmlAssist.GetTagValue(objCustomerRoleElem, "APPLICATIONNUMBER")
    strCustomerNo = m_objXmlAssist.GetTagValue(objCustomerRoleElem, "CUSTOMERNUMBER")

    If m_objContext Is Nothing Then
        Set objIApplManagerDO = New ApplicationManagerDO
        Set objCustomerBO = New CustomerBO
    Else
        Set objIApplManagerDO = m_objContext.CreateInstance(cstrCOMPONENT_NAME & _
                                                             ".ApplicationManagerDO")
        Set objCustomerBO = m_objContext.CreateInstance(gstrCUSTOMER_COMPONENT & ".CustomerBO")
    End If

    ' Set up delete request for CustomerLockApplicationLock
    Set objRequestNode = m_objXmlAssist.GetRequestNode(objCustomerRoleElem.ownerDocument)
    objRequestDoc.appendChild objRequestNode
    Set objTableElem = objRequestDoc.createElement("CUSTOMERLOCKAPPLICATIONLOCK")
    objRequestNode.appendChild objTableElem
    Set objElement = objRequestDoc.createElement("APPLICATIONNUMBER")
    objElement.Text = strApplicationNo
    objTableElem.appendChild objElement
    Set objElement = objRequestDoc.createElement("CUSTOMERNUMBER")
    objElement.Text = strCustomerNo
    objTableElem.appendChild objElement

    objIApplManagerDO.DeleteCustomerLockApplicationLock objRequestNode
    
    ' Delete Customer Lock
    Dim strResponse As String
    Set objRequestNode = objRequestDoc.createElement("REQUEST")
    Set objTableElem = objRequestDoc.createElement("CUSTOMERLOCK")
    objRequestNode.appendChild objTableElem

    Set objElement = objRequestDoc.createElement("CUSTOMERNUMBER")
    objElement.Text = strCustomerNo
    objTableElem.appendChild objElement

    strResponse = objCustomerBO.DeleteCustomerLock(objRequestNode.xml)
    m_objErrAssist.CheckXMLResponse strResponse, True

    Set objRequestDoc = Nothing
    Set objTableElem = Nothing
    Set objRequestNode = Nothing
    Set objElement = Nothing
    Set objIApplManagerDO = Nothing

    Exit Sub

DeleteCustomerLocksVbErr:

    Set objRequestDoc = Nothing
    Set objTableElem = Nothing
    Set objRequestNode = Nothing
    Set objElement = Nothing
    Set objIApplManagerDO = Nothing

    ' Re-raise error for business object to interpret as appropriate
    Err.Raise Err.Number, Err.Source, Err.Description

End Sub

Private Function IApplicationManagerTxBO_DeleteCustomerRole( _
                    ByVal vxmlRequest As IXMLDOMElement) As IXMLDOMNode
' header ----------------------------------------------------------------------------------
' description:  Delete multiple instances of the persistant data associated with the table
'               'CustomerRole'
'               invokes appropriate Data Object(s) to perform persistance
' pass:         vxmlRequest  xml Request element
' return:       xml Response Node
'------------------------------------------------------------------------------------------
On Error GoTo DeleteVbErr

    Const strFunctionName As String = "IApplicationManagerTxBO_DeleteCustomerRole"

    Dim xmlNodeList As IXMLDOMNodeList
    Dim xmlElement As IXMLDOMElement
    Dim xmlOut As New FreeThreadedDOMDocument40
    xmlOut.validateOnParse = False
    xmlOut.setProperty "NewParser", True
    Dim xmlResponseElem As IXMLDOMElement
   
    Dim objIApplManagerDO As IApplicationManagerDO
    
    Set xmlResponseElem = xmlOut.createElement("RESPONSE")
    xmlOut.appendChild xmlResponseElem
    xmlResponseElem.setAttribute "TYPE", "SUCCESS"
                
    ' Check if the element passed in is the one with the table name else
    ' look below the node that is passed in
    If vxmlRequest.nodeName = cstTABLE_CUSTOMERROLE Then
        Set xmlElement = vxmlRequest
    Else
        Set xmlNodeList = vxmlRequest.getElementsByTagName(cstTABLE_CUSTOMERROLE)
    End If
    
    If xmlElement Is Nothing Then
        If xmlNodeList.length = 0 Then
            m_objErrAssist.ThrowError TypeName(Me), _
                                  strFunctionName, _
                                  oeMissingPrimaryTag, _
                                  cstTABLE_CUSTOMERROLE & " tag not found"
        End If
    End If
   
    If m_objContext Is Nothing Then
        Set objIApplManagerDO = New ApplicationManagerDO
    Else
        Set objIApplManagerDO = m_objContext.CreateInstance(App.Title & ".ApplicationManagerDO")
    End If
    
    If Not xmlElement Is Nothing Then
        objIApplManagerDO.DeleteCustomerRole xmlElement
        DeleteCustomerLocks xmlElement
    Else
        For Each xmlElement In xmlNodeList
            objIApplManagerDO.DeleteCustomerRole xmlElement
            DeleteCustomerLocks xmlElement
        Next xmlElement
    End If
    
    Set IApplicationManagerTxBO_DeleteCustomerRole = xmlResponseElem
    
    If Not m_objContext Is Nothing Then
        m_objContext.SetComplete
    End If
    
DeleteExit:

    Set objIApplManagerDO = Nothing
    Set xmlNodeList = Nothing
    Set xmlElement = Nothing
    Set xmlOut = Nothing
    Set xmlResponseElem = Nothing
        
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

    Set IApplicationManagerTxBO_DeleteCustomerRole = m_objErrAssist.CreateErrorResponseEx()

    If Not m_objContext Is Nothing Then
        m_objContext.SetAbort
    End If
    
    ' go to clean-up section
    Resume DeleteExit

End Function

Private Function IApplicationManagerTxBO_DeleteLock( _
            ByVal vxmlRequest As IXMLDOMElement) As IXMLDOMNode
' header ----------------------------------------------------------------------------------
' description:  Delete multiple instances of the persistant data associated with the table
'               'ApplicationLock'
'               invokes appropriate Data Object(s) to perform persistance
' pass:         vxmlRequest  xml Request element
' return:       xml Response Node
'------------------------------------------------------------------------------------------
On Error GoTo DeleteLockVbErr

    Const strFunctionName As String = "IApplicationManagerTxBO_DeleteLock"

    Dim xmlElement As IXMLDOMElement
    Dim xmlOut As New FreeThreadedDOMDocument40
    xmlOut.validateOnParse = False
    xmlOut.setProperty "NewParser", True
    Dim xmlResponseElem As IXMLDOMElement
    
    Dim xmlRequestNode As IXMLDOMNode
    Dim xmlTempElem As IXMLDOMElement
    Dim xmlAccessAuditTableNode As IXMLDOMNode

    Dim objIApplicationManagerDO As IApplicationManagerDO
    Dim objIAuditDO As IAuditDO
    Dim objAuditBO As AuditBO
    
    Set xmlResponseElem = xmlOut.createElement("RESPONSE")
    xmlOut.appendChild xmlResponseElem
    xmlResponseElem.setAttribute "TYPE", "SUCCESS"
                
    ' Check if the element passed in is the one with the table name else
    ' look below the node that is passed in
    If vxmlRequest.nodeName = cstrTABLE_APPLLOCK Then
        Set xmlElement = vxmlRequest
    Else
        Set xmlElement = vxmlRequest.getElementsByTagName(cstrTABLE_APPLLOCK).Item(0)
    End If
    
    If xmlElement Is Nothing Then
        m_objErrAssist.ThrowError TypeName(Me), _
                                  strFunctionName, _
                                  oeMissingPrimaryTag, _
                                  cstrTABLE_APPLLOCK & " tag not found"
    End If
   
    If m_objContext Is Nothing Then
        Set objIApplicationManagerDO = New ApplicationManagerDO
        Set objIAuditDO = New AuditDO
        Set objAuditBO = New AuditBO
    Else
        Set objIApplicationManagerDO = m_objContext.CreateInstance(App.Title & ".ApplicationManagerDO")
        Set objIAuditDO = _
            m_objContext.CreateInstance(gstrAUDIT_COMPONENT & ".AuditDO")
        Set objAuditBO = m_objContext.CreateInstance(gstrAUDIT_COMPONENT & ".AuditBO")
    End If
    
    objIApplicationManagerDO.DeleteLock xmlElement
    
    '------------------------------------------------------------------------------------------
    ' RF 29/09/99
    ' On deletion of an application lock, write an audit record.
    '------------------------------------------------------------------------------------------
    
    Set xmlRequestNode = vxmlRequest.cloneNode(True)
        
    m_objXmlAssist.ChangeNodeName xmlRequestNode, "APPLICATIONLOCK", "ACCESSAUDIT"
    Set xmlAccessAuditTableNode = xmlRequestNode.selectSingleNode(".//ACCESSAUDIT")
        
    Set xmlTempElem = xmlOut.createElement("AUDITRECORDTYPE")
    xmlTempElem.Text = objIAuditDO.GetApplicationReleaseValueId()
    xmlAccessAuditTableNode.appendChild xmlTempElem
        
    Set xmlTempElem = xmlOut.createElement("SUCCESSINDICATOR")
    xmlTempElem.Text = "1"
    xmlAccessAuditTableNode.appendChild xmlTempElem
     
    Dim strResponse As String
    strResponse = objAuditBO.CreateAccessAudit(xmlRequestNode.xml)
    m_objErrAssist.CheckXMLResponse strResponse, True
    
    Set IApplicationManagerTxBO_DeleteLock = xmlResponseElem
    
    If Not m_objContext Is Nothing Then
        m_objContext.SetComplete
    End If
    
DeleteLockExit:

    Set objIApplicationManagerDO = Nothing
    Set xmlElement = Nothing
    Set xmlOut = Nothing
    Set xmlResponseElem = Nothing
        
    Set xmlRequestNode = Nothing
    Set xmlTempElem = Nothing
    Set xmlAccessAuditTableNode = Nothing
    
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

    Set IApplicationManagerTxBO_DeleteLock = m_objErrAssist.CreateErrorResponseEx()

    If Not m_objContext Is Nothing Then
        m_objContext.SetAbort
    End If
    
    ' go to clean-up section
    Resume DeleteLockExit

End Function

Private Function IApplicationManagerTxBO_ImportAccountRefreshIntoApplication( _
                                ByVal vxmlRequest As IXMLDOMElement) As IXMLDOMNode

On Error GoTo ImportAccountRefreshIntoApplicationVbErr

Const cstrFunctionName As String = "IApplicationManagerTxBO_ImportAccountRefreshIntoApplication"

    Dim strCustomerNumber As String, strOtherSysCustomerNumber As String
    Dim strResponse As String, strAccountNumber As String
    Dim strApplicationNumber As String, strAFFNumber As String, strPortingIndicator As String
    Dim strBMAccountNumber As String, strMortgageLoanGuid As String, strAccountGUID As String
    Dim intCount As Integer
    Dim strICBSTerm As String, strFirstAcc As String
        
    Dim xmlOut As FreeThreadedDOMDocument40, xmlResponseElem As IXMLDOMElement
    Dim xmlApplication As IXMLDOMNode, xmlNewRequest As IXMLDOMElement
    Dim xmlTable As IXMLDOMNode, xmlElem1 As IXMLDOMNode, xmlNode As IXMLDOMNode
    Dim xmlTableList As IXMLDOMNode, xmlAdminLoan As IXMLDOMNode
    Dim xmlCustomerList As IXMLDOMNodeList, xmlAccountRefresh As IXMLDOMNode
    Dim xmlMortAccAndLoan As IXMLDOMNode, xmlMortgageLoanList As IXMLDOMNodeList
        
    Dim objCFBO As CustomerFinancialBO, objIApplicationDO As IApplicationDO
    Dim objAdminBO As Object
        
    Set xmlOut = New FreeThreadedDOMDocument40
    xmlOut.validateOnParse = False
    xmlOut.setProperty "NewParser", True
    Set xmlResponseElem = xmlOut.createElement("RESPONSE")
    xmlOut.appendChild xmlResponseElem
    xmlResponseElem.setAttribute "TYPE", "SUCCESS"
        
    Set objAdminBO = m_objContext.CreateInstance(gstrADMIN_INTERFACE & ".AdminInterfaceBO")
    Set objCFBO = m_objContext.CreateInstance(gstrCUSTOMER_FINANCIAL & ".CustomerFinancialBO")
    Set objIApplicationDO = m_objContext.CreateInstance(App.Title & ".ApplicationDO")
    
    '--------------------------------------------------------------------------
    'Call AccountRefresh (of omAdminBO)
    '--------------------------------------------------------------------------
    Set xmlNewRequest = m_objXmlAssist.GetRequestNodeEx(vxmlRequest)
    xmlNewRequest.setAttribute "OPERATION", "GETACCOUNTREFRESH"
    Set xmlTable = xmlOut.createElement("IMPORTACCOUNTREFRESHINTOAPPLICATION")
    xmlNewRequest.appendChild xmlTable
    
    Set xmlApplication = vxmlRequest.selectSingleNode(".//APPLICATION")
    Set xmlElem1 = xmlOut.createElement("APPLICATION")
    xmlTable.appendChild xmlElem1
    xmlSetAttributeValue xmlElem1, "MORTGAGESTATUSTYPE", _
                m_objXmlAssist.GetTagValue(xmlApplication, "TYPEOFAPPLICATION")
    strBMAccountNumber = m_objXmlAssist.GetTagValue(xmlApplication, "BMACCOUNTNUMBER")
                
    Set xmlCustomerList = vxmlRequest.selectNodes(".//CUSTOMER")
    If xmlCustomerList.length > 0 Then
        Set xmlTableList = xmlOut.createElement("CUSTOMERLIST")
        xmlTable.appendChild xmlTableList
    End If
    For Each xmlNode In xmlCustomerList
        Set xmlElem1 = xmlOut.createElement("CUSTOMER")
        xmlSetAttributeValue xmlElem1, "OMIGACUSTOMERNUMBER", m_objXmlAssist.GetTagValue(xmlNode, "CUSTOMERNUMBER")
        xmlSetAttributeValue xmlElem1, "CUSTOMERNUMBER", m_objXmlAssist.GetTagValue(xmlNode, "OTHERSYSTEMCUSTOMERNUMBER")
        xmlSetAttributeValue xmlElem1, "BMACCOUNTNUMBER", strBMAccountNumber
        
        xmlTableList.appendChild xmlElem1
    Next xmlNode
    
    strResponse = objAdminBO.REQUEST(xmlNewRequest.xml)
    Set xmlOut = m_objXmlAssist.load(strResponse, TypeName(Me), cstrFunctionName)
    m_objErrAssist.CheckResponse xmlOut.documentElement, xmlResponseElem, True
    
    Set xmlAccountRefresh = xmlOut.documentElement
    
    '--------------------------------------------------------------------------
    'Get MortgageAccount and MortgageLoan details for the customers passed in
    '--------------------------------------------------------------------------
    Set xmlNewRequest = m_objXmlAssist.GetRequestNodeEx(vxmlRequest)
    Set xmlTableList = xmlOut.createElement("MORTGAGEACCOUNTANDLOANLIST")
    xmlNewRequest.appendChild xmlTableList
    
    Set xmlCustomerList = vxmlRequest.selectNodes(".//CUSTOMER")
    For Each xmlNode In xmlCustomerList
        Set xmlTable = xmlOut.createElement("MORTGAGEACCOUNTANDLOAN")
        Set xmlElem1 = xmlOut.createElement("CUSTOMERNUMBER")
        xmlElem1.Text = m_objXmlAssist.GetTagValue(xmlNode, "CUSTOMERNUMBER")
        xmlTable.appendChild xmlElem1
        
        Set xmlElem1 = xmlOut.createElement("CUSTOMERVERSIONNUMBER")
        xmlElem1.Text = m_objXmlAssist.GetTagValue(xmlNode, "CUSTOMERVERSIONNUMBER")
        xmlTable.appendChild xmlElem1
        
        xmlTableList.appendChild xmlTable
    Next xmlNode
    
    strResponse = objCFBO.FindMortgageAccountAndLoanList(xmlNewRequest.xml)
    Set xmlOut = m_objXmlAssist.load(strResponse, TypeName(Me), cstrFunctionName)
    m_objErrAssist.CheckResponse xmlOut.documentElement, xmlResponseElem, True
    Set xmlMortAccAndLoan = xmlOut.documentElement
    
    '-----------------------------------------------------------------------------------------
    'Update Mortgage Loans - taking data from output of GetAccountRefresh
    '-----------------------------------------------------------------------------------------
    strPortingIndicator = m_objXmlAssist.GetTagValue(vxmlRequest, ".//PORTINGINDICATOR")
    
    Set xmlNewRequest = m_objXmlAssist.GetRequestNodeEx(vxmlRequest)
    Set xmlTableList = xmlOut.createElement("MORTGAGELOANLIST")
    xmlNewRequest.appendChild xmlTableList
    
    Set xmlMortgageLoanList = xmlMortAccAndLoan.selectNodes(".//MORTGAGELOAN")
    intCount = 0
    strICBSTerm = "0"
    For Each xmlNode In xmlMortgageLoanList
        strMortgageLoanGuid = xmlNode.selectSingleNode("MORTGAGELOANGUID").Text
        strAccountGUID = xmlNode.selectSingleNode("ACCOUNTGUID").Text
        'Same Loan GUID can appear in mutliple records. so check whether it was already added
        'to the UPDATE request
        Set xmlElem1 = xmlNewRequest.selectSingleNode(".//MORTGAGELOAN[MORTGAGELOANGUID='" & strMortgageLoanGuid & "']")
        If xmlElem1 Is Nothing Then
       
            strAccountNumber = m_objXmlAssist.GetTagValue(xmlNode, "LOANACCOUNTNUMBER")
            Set xmlAdminLoan = xmlAccountRefresh.selectSingleNode( _
                                    ".//MORTGAGELOAN[@LOANACCOUNTNUMBER='" & strAccountNumber & "']")
            If Not xmlAdminLoan Is Nothing Then
                intCount = intCount + 1
                Set xmlTable = xmlOut.createElement("MORTGAGELOAN")
                xmlTableList.appendChild xmlTable
                
                Set xmlElem1 = xmlOut.createElement("ACCOUNTGUID")
                xmlElem1.Text = strAccountGUID
                xmlTable.appendChild xmlElem1
                
                Set xmlElem1 = xmlOut.createElement("MORTGAGELOANGUID")
                xmlElem1.Text = strMortgageLoanGuid
                xmlTable.appendChild xmlElem1
                
                Set xmlElem1 = xmlOut.createElement("PENALTYPLANCODE")
                xmlElem1.Text = m_objXmlAssist.GetAttributeFromNode(xmlAdminLoan, "PENALTYPLANCODE")
                xmlTable.appendChild xmlElem1
                
                Set xmlElem1 = xmlOut.createElement("PRODUCTSTARTDATE")
                xmlElem1.Text = m_objXmlAssist.GetAttributeFromNode(xmlAdminLoan, "PRODUCTSTARTDATE")
                xmlTable.appendChild xmlElem1
                
                Set xmlElem1 = xmlOut.createElement("PRODUCTSTEP")
                xmlElem1.Text = m_objXmlAssist.GetAttributeFromNode(xmlAdminLoan, "PRODUCTSTEP")
                xmlTable.appendChild xmlElem1
                
                Set xmlElem1 = xmlOut.createElement("REMAININGSTEPDURATION")
                xmlElem1.Text = m_objXmlAssist.GetAttributeFromNode(xmlAdminLoan, "REMAININGSTEPDURATION")
                xmlTable.appendChild xmlElem1
                
                If strPortingIndicator <> 1 Then
                    Set xmlElem1 = xmlOut.createElement("OUTSTANDINGBALANCE")
                    xmlElem1.Text = m_objXmlAssist.GetAttributeFromNode(xmlAdminLoan, "OUTSTANDINGBALANCE")
                    xmlTable.appendChild xmlElem1
                    
                    Set xmlElem1 = xmlOut.createElement("REMAININGINTERESTONLYAMOUNT")
                    xmlElem1.Text = m_objXmlAssist.GetAttributeFromNode(xmlAdminLoan, "REMAININGINTERESTONLYAMOUNT")
                    xmlTable.appendChild xmlElem1
                    
                    Set xmlElem1 = xmlOut.createElement("REMAININGCAPITALINTERESTAMOUNT")
                    xmlElem1.Text = m_objXmlAssist.GetAttributeFromNode(xmlAdminLoan, "REMAININGCAPITALINTERESTAMOUNT")
                    xmlTable.appendChild xmlElem1
                    
                    'JD BMIDS975 Update the term if the mortgage validation type is ET
                    If IsValidationType("TypeOfMortgage", m_objXmlAssist.GetTagValue(xmlApplication, "TYPEOFAPPLICATION"), "ET") Then
                        Set xmlElem1 = xmlOut.createElement("ICBSCALCULATEDOUTSTANDINGTERM")
                        xmlElem1.Text = m_objXmlAssist.GetAttributeFromNode(xmlAdminLoan, "ICBSCALCULATEDOUTSTANDINGTERM")
                        xmlTable.appendChild xmlElem1
                        'save the term if it is the one for the first mortgageaccount
                        strFirstAcc = m_objXmlAssist.GetTagValue(xmlApplication, "BMACCOUNTNUMBER") + "/1"
                        If m_objXmlAssist.GetTagValue(xmlNode, "LOANACCOUNTNUMBER") = strFirstAcc Then
                            strICBSTerm = m_objXmlAssist.GetAttributeFromNode(xmlAdminLoan, "ICBSCALCULATEDOUTSTANDINGTERM")
                        End If
                    End If
                End If
            End If
        End If
    Next xmlNode
    
    If intCount > 0 Then
        strResponse = objCFBO.UpdateMortgageLoan(xmlNewRequest.xml)
        Set xmlOut = m_objXmlAssist.load(strResponse, TypeName(Me), cstrFunctionName)
        m_objErrAssist.CheckResponse xmlOut.documentElement, xmlResponseElem, True
    End If
    
    '-------------------------------------------------------------------------------------
    'Update Application (Column --> LastAccountDownLoadDate)
    '-------------------------------------------------------------------------------------
    strApplicationNumber = m_objXmlAssist.GetTagValue(xmlApplication, "APPLICATIONNUMBER")
    strAFFNumber = m_objXmlAssist.GetTagValue(xmlApplication, "APPLICATIONFACTFINDNUMBER")
    
    Set xmlTable = xmlOut.createElement("APPLICATION")
    Set xmlElem1 = xmlOut.createElement("APPLICATIONNUMBER")
    xmlElem1.Text = strApplicationNumber
    xmlTable.appendChild xmlElem1
    
    Set xmlElem1 = xmlOut.createElement("APPLICATIONFACTFINDNUMBER")
    xmlElem1.Text = strApplicationNumber
    xmlTable.appendChild xmlElem1
    
    Set xmlElem1 = xmlOut.createElement("LASTACCOUNTDOWNLOADDATE")
    xmlElem1.Text = Now()
    xmlTable.appendChild xmlElem1
    
    objIApplicationDO.Update xmlTable
    
    'JD  BMIDS975 return the ICBSOutstandingTerm
    Set xmlElem1 = xmlOut.createElement("TERM")
    Set xmlNode = xmlResponseElem.appendChild(xmlElem1)
    Set xmlElem1 = xmlOut.createElement("ICBSOUTSTANDINGTERM")
    xmlElem1.Text = strICBSTerm
    xmlNode.appendChild xmlElem1
    
    Set IApplicationManagerTxBO_ImportAccountRefreshIntoApplication = xmlResponseElem
    
ImportAccountRefreshIntoApplicationExit:

    Set xmlOut = Nothing
    Set xmlResponseElem = Nothing
    Set xmlApplication = Nothing
    Set xmlNewRequest = Nothing
    Set xmlTableList = Nothing
    Set xmlAdminLoan = Nothing
    Set xmlTable = Nothing
    Set xmlElem1 = Nothing
    Set xmlNode = Nothing
    Set xmlAdminLoan = Nothing
    Set xmlCustomerList = Nothing
    Set xmlAccountRefresh = Nothing
        
    Set objCFBO = Nothing
    Set objAdminBO = Nothing
    
    Exit Function
ImportAccountRefreshIntoApplicationVbErr:
   
    If m_objErrAssist.IsWarning = True Then
        m_objErrAssist.AddWarning xmlResponseElem
        Resume Next
    End If

    m_objErrAssist.AddToErrSource cstrFunctionName
    
    If m_objErrAssist.IsSystemError = True Then
        m_objErrAssist.LogError TypeName(Me), cstrFunctionName, Err.Number, Err.Description
    End If

    Set IApplicationManagerTxBO_ImportAccountRefreshIntoApplication = m_objErrAssist.CreateErrorResponseEx()

    m_objContext.SetAbort
    
    ' go to clean-up section
    Resume ImportAccountRefreshIntoApplicationExit
    
End Function

Private Function IApplicationManagerTxBO_ImportCustomersIntoApplication(ByVal vxmlRequest As IXMLDOMElement) As IXMLDOMNode
' header ----------------------------------------------------------------------------------
' procedure:    IApplicationManagerTxBO_ImportCustomersIntoApplication
' author:       Mark Coleman
' date:         17/07/2002
' description:
' pass:         vxmlRequestNode  XML Input request specifying the data and operation
' return:       vxmlResponseNode XML Response
'------------------------------------------------------------------------------------------
On Error GoTo ImportCustomersIntoApplicationVbErr

Const cstrFunctionName As String = "IApplicationManagerTxBO_ImportCustomersIntoApplication"
    
Dim xmlElement As IXMLDOMElement
Dim xmlOut As FreeThreadedDOMDocument40
Dim xmlResponseElem As IXMLDOMElement
Dim objAdminBO As Object
Dim objComboDO As ComboDO
Dim xmlAdminRequestElem As IXMLDOMElement
Dim xmlCustomerElem As IXMLDOMElement
Dim xmlAdminResponseDoc As FreeThreadedDOMDocument40
Dim objIApplicationBO As IApplicationBO
Dim xmlAppNo As IXMLDOMElement
Dim xmlAppFFNo As IXMLDOMElement
Dim xmlTempRequestDoc As FreeThreadedDOMDocument40
Dim xmlRequestElem As IXMLDOMElement
Dim xmlTableElem As IXMLDOMElement
Dim xmlTempElement As IXMLDOMElement
Dim objConvAssist As ConvertAssist
Dim xmlRequestCopy As IXMLDOMElement
Dim xmlElemBasedAdmin As IXMLDOMNode
Dim xmlCustomerList As IXMLDOMNodeList

Dim strInCustomerNo As String
Dim strInCustVersNo As String
Dim strAccountNo As String
Dim strTypeOfApplication As String
Dim blnFurtherAdvance As Boolean
Dim strAdminResponse As String
Dim strPattern As String
Dim strOtherSysCustNo As String
Dim blnAddAllCustomers As Boolean
Dim strValueId As String    'BMIDS00527 MDC 24/10/2002
        
Const cstrTABLE_ACCOUNT = "ACCOUNT"
        
        
    'Set up Response
    Set xmlOut = New FreeThreadedDOMDocument40
    xmlOut.validateOnParse = False
    xmlOut.setProperty "NewParser", True
    Set xmlResponseElem = xmlOut.createElement("RESPONSE")
    xmlOut.appendChild xmlResponseElem
    xmlResponseElem.setAttribute "TYPE", "SUCCESS"
                
    'Validate Request
    If vxmlRequest.nodeName = cstrTABLE_ACCOUNT Then
        Set xmlElement = vxmlRequest
    Else
        Set xmlElement = vxmlRequest.getElementsByTagName(cstrTABLE_ACCOUNT).Item(0)
    End If
    If xmlElement Is Nothing Then
        m_objErrAssist.ThrowError TypeName(Me), _
                                  cstrFunctionName, _
                                  oeMissingPrimaryTag, _
                                  cstrTABLE_ACCOUNT & " tag not found"
    End If
    
    
    'Initialise
    Set objConvAssist = New ConvertAssist
    Set objComboDO = m_objContext.CreateInstance(gstrBASE_COMPONENT & ".ComboDO")
    Set objIApplicationBO = m_objContext.CreateInstance(App.Title & ".ApplicationBO")
    Set objAdminBO = m_objContext.CreateInstance(gstrADMIN_INTERFACE & ".AdminInterfaceBO")
    strInCustomerNo = m_objXmlAssist.GetNodeValue(vxmlRequest, ".//CUSTOMERNUMBER")
'    strInOtherSysCustNo = m_objXmlAssist.GetNodeValue(xmlElement, ".//OTHERSYSTEMCUSTOMERNUMBER")
    strInCustVersNo = m_objXmlAssist.GetNodeValue(vxmlRequest, ".//CUSTOMERVERSIONNUMBER")
    strTypeOfApplication = m_objXmlAssist.GetNodeValue(vxmlRequest, ".//TYPEOFAPPLICATION")
    strAccountNo = m_objXmlAssist.GetNodeValue(xmlElement, ".//ACCOUNTNUMBER", False, False)
'    strAddAllCustomers = m_objXmlAssist.GetNodeValue(xmlElement, ".//ADDALLCUSTOMERS", False, False)
'    blnAddAllCustomers = objConvAssist.CSafeBool(strAddAllCustomers)
    blnFurtherAdvance = objComboDO.IsItemInValidation("TypeOfMortgage", strTypeOfApplication, "F")
    If blnFurtherAdvance = True And Len(strAccountNo) = 0 Then
        m_objErrAssist.ThrowError TypeName(Me), _
                                  cstrFunctionName, _
                                  oeMissingElementValue, _
                                  "Further advance requires account number"
    End If
    
    '========================================================================================================
    'Get details for all Customers on the requested account
    '========================================================================================================
    
    'Set up Request for call to AdminInterfaceBO method GetAccountCustomers
    Set xmlAdminRequestElem = xmlCreateAttributeBasedResponse(vxmlRequest, True)
    'BMIDS00475 Format XML so that it conforms to GetAccountCustomers request DTD
    m_objXmlAssist.PromoteNode xmlAdminRequestElem, m_objXmlAssist.GetMandatoryNode(xmlAdminRequestElem, "IMPORTCUSTOMERSINTOAPPLICATION/ACCOUNT"), True, False
    'BMIDS00475 End
    xmlAdminRequestElem.setAttribute "OPERATION", "GetAccountCustomers"
    strAdminResponse = objAdminBO.REQUEST(xmlAdminRequestElem.xml)
    Set xmlAdminResponseDoc = m_objXmlAssist.load(strAdminResponse, TypeName(Me), cstrFunctionName)
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
    
    
    '========================================================================================================
    'Create and lock a new application
    '========================================================================================================
    'Set up the request
    Set xmlRequestCopy = vxmlRequest.cloneNode(True)
    'xmlRequestCopy.removeChild xmlRequestCopy.selectSingleNode(".//CUSTOMER")
    Set xmlTableElem = m_objXmlAssist.GetNode(xmlRequestCopy, ".//APPLICATION")
    Set xmlTempElement = vxmlRequest.ownerDocument.createElement("UNITID")
    xmlTempElement.Text = m_objXmlAssist.GetAttributeFromNode(vxmlRequest, "UNITID")
    xmlTableElem.appendChild xmlTempElement
    Set xmlTempElement = vxmlRequest.ownerDocument.createElement("USERID")
    xmlTempElement.Text = m_objXmlAssist.GetAttributeFromNode(vxmlRequest, "USERID")
    xmlTableElem.appendChild xmlTempElement
    Set xmlTempElement = vxmlRequest.ownerDocument.createElement("CHANNELID")
    xmlTempElement.Text = m_objXmlAssist.GetAttributeFromNode(vxmlRequest, "CHANNELID")
    xmlTableElem.appendChild xmlTempElement
    
    ' Create the application
    Set xmlTempElement = objIApplicationBO.CreateApplication(xmlRequestCopy)
    m_objErrAssist.CheckResponse xmlTempElement, xmlResponseElem, True
    Set xmlAppNo = m_objXmlAssist.GetNode(xmlTempElement, ".//APPLICATIONNUMBER")
    Set xmlAppFFNo = m_objXmlAssist.GetNode(xmlTempElement, ".//APPLICATIONFACTFINDNUMBER")
    Set xmlTempElement = Nothing
    
    ' Append the generated keys to the response
    Set xmlTempElement = xmlOut.createElement("APPLICATION")
    xmlTempElement.appendChild xmlAppNo.cloneNode(True)
    xmlTempElement.appendChild xmlAppFFNo.cloneNode(True)
    xmlResponseElem.appendChild xmlTempElement
    Set xmlTempElement = Nothing
    
    
    '========================================================================================================
    'Create Financial Summary
    '========================================================================================================
    
    Set xmlTempRequestDoc = New FreeThreadedDOMDocument40
    xmlTempRequestDoc.validateOnParse = False
    xmlTempRequestDoc.setProperty "NewParser", True
    xmlTempRequestDoc.async = False
    
    ' Create Financial Summary
    Set xmlRequestElem = m_objXmlAssist.GetRequestNodeEx(vxmlRequest)
    xmlTempRequestDoc.appendChild xmlRequestElem
    
    Set xmlTableElem = xmlTempRequestDoc.createElement("FINANCIALSUMMARY")
    xmlRequestElem.appendChild xmlTableElem
    xmlTableElem.appendChild xmlAppNo.cloneNode(True)
    xmlTableElem.appendChild xmlAppFFNo.cloneNode(True)
    
    'Existing Mortgage
    Set xmlTempElement = xmlTempRequestDoc.createElement("EXISTINGMORTGAGEINDICATOR")
    xmlTempElement.Text = "1"
'    If xmlAccountList.length > 0 Then
'        xmlTempElement.Text = "1"
'    Else
'        xmlTempElement.Text = "0"
'    End If
    xmlTableElem.appendChild xmlTempElement
    
    'Arrears History
'    Set xmlTempElement = xmlTempRequestDoc.createElement("ARREARSHISTORYINDICATOR")
'    If Len(strMaxBalance) > 0 Then
'        xmlTempElement.Text = "1"
'    Else
'        xmlTempElement.Text = "0"
'    End If
'    xmlTableElem.appendChild xmlTempElement
    
    Set xmlTempElement = objIApplicationBO.CreateFinancialSummary(xmlRequestElem)
    m_objErrAssist.CheckResponse xmlTempElement, xmlResponseElem, True
    
    
    '========================================================================================================
    'Create Customer records as required and link them to the application
    '========================================================================================================
    'Determine if the current customer is the only one on the account.
    strOtherSysCustNo = m_objXmlAssist.GetNodeValue(vxmlRequest, ".//OTHERSYSTEMCUSTOMERNUMBER")
    strPattern = ".//CUSTOMER[CUSTOMERNUMBER != '" & strOtherSysCustNo & "']"
    Set xmlCustomerList = xmlElemBasedAdmin.selectNodes(strPattern)
    If xmlCustomerList.length > 0 Then
        blnAddAllCustomers = True
    End If
    
    Call ImportCustomers(vxmlRequest, xmlElemBasedAdmin, xmlAppNo.Text, xmlAppFFNo.Text, _
                                    True, blnAddAllCustomers, xmlResponseElem, "")
    
    'BMIDS00527 MDC 23/10/2002 - Update Application with Number of Applicants & Guarantors
    Set xmlRequestElem = m_objXmlAssist.GetRequestNodeEx(vxmlRequest)
    Set xmlTableElem = vxmlRequest.ownerDocument.createElement("APPLICATIONFACTFIND")
    xmlRequestElem.appendChild xmlTableElem
    xmlTableElem.appendChild xmlAppNo
    xmlTableElem.appendChild xmlAppFFNo
    Set xmlTempElement = vxmlRequest.ownerDocument.createElement("NUMBEROFAPPLICANTS")
    strValueId = objComboDO.GetFirstComboValueId("CustomerRoleType", "A")   'Applicant
    xmlTempElement.Text = xmlElemBasedAdmin.selectNodes(".//CUSTOMER[CUSTOMERROLETYPE='" & strValueId & "']").length
    xmlTableElem.appendChild xmlTempElement
    Set xmlTempElement = vxmlRequest.ownerDocument.createElement("NUMBEROFGUARANTORS")
    strValueId = objComboDO.GetFirstComboValueId("CustomerRoleType", "G")   'Guarantor
    xmlTempElement.Text = xmlElemBasedAdmin.selectNodes(".//CUSTOMER[CUSTOMERROLETYPE='" & strValueId & "']").length
    xmlTableElem.appendChild xmlTempElement
        
    Set xmlTempElement = objIApplicationBO.UpdateApplicationFactFind(xmlRequestElem)
    m_objErrAssist.CheckResponse xmlTempElement, xmlResponseElem, True
    'BMIDS00527 MDC 23/10/2002 - End
        
    'Return application keys
    Set IApplicationManagerTxBO_ImportCustomersIntoApplication = xmlResponseElem
    m_objContext.SetComplete
    
ImportCustomersIntoApplicationExit:

    Set xmlElement = Nothing
    Set xmlOut = Nothing
    Set xmlResponseElem = Nothing
    Set objAdminBO = Nothing
    Set objComboDO = Nothing
    Set xmlAdminRequestElem = Nothing
    Set xmlCustomerElem = Nothing
    Set xmlAdminResponseDoc = Nothing
    Set objIApplicationBO = Nothing
    Set xmlAppNo = Nothing
    Set xmlAppFFNo = Nothing
    Set xmlTempRequestDoc = Nothing
    Set xmlRequestElem = Nothing
    Set xmlTableElem = Nothing
    Set xmlTempElement = Nothing
    Set objConvAssist = Nothing
    Set xmlRequestCopy = Nothing
    Set xmlElemBasedAdmin = Nothing
    Set xmlCustomerList = Nothing
    
    
    Exit Function

ImportCustomersIntoApplicationVbErr:

    If m_objErrAssist.GetOmigaErrorNumber(Err.Number) = oeRecordNotFound Then
        Resume Next
    End If
    
    If m_objErrAssist.IsWarning = True Then
        m_objErrAssist.AddWarning xmlResponseElem
        Resume Next
    End If

    m_objErrAssist.AddToErrSource cstrFunctionName
    
    If m_objErrAssist.IsSystemError = True Then
        m_objErrAssist.LogError TypeName(Me), cstrFunctionName, Err.Number, Err.Description
    End If

    Set IApplicationManagerTxBO_ImportCustomersIntoApplication = m_objErrAssist.CreateErrorResponseEx()

    m_objContext.SetAbort
    
    ' go to clean-up section
    Resume ImportCustomersIntoApplicationExit

End Function

Private Function IApplicationManagerTxBO_LockCustomersForApplication( _
                    ByVal vxmlRequest As IXMLDOMElement) As IXMLDOMNode
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
' return:   n/a
'------------------------------------------------------------------------------------------
On Error GoTo LockCustomersForApplicationVbErr
    
    Const strFunctionName As String = "IApplicationManagerTxBO_LockCustomersForApplication"
    
    Dim xmlResponseElem As IXMLDOMElement
    Dim xmlTempResponseNode As IXMLDOMElement
    Dim xmlOut As New FreeThreadedDOMDocument40
    xmlOut.validateOnParse = False
    xmlOut.setProperty "NewParser", True
    Dim objIApplManagerDO As IApplicationManagerDO
    Dim objAuditBO As AuditBO, objIAuditDO As IAuditDO 'SR MAR1631
    Dim xmlRequestElem As IXMLDOMElement, xmlAccessAudit As IXMLDOMElement, xmlTempElem As IXMLDOMElement 'SR MAR1631
    Dim strResponse As String 'SR MAR1631
   
    'Create a default response element.
    Set xmlResponseElem = xmlOut.createElement("RESPONSE")
    xmlOut.appendChild xmlResponseElem
    xmlResponseElem.setAttribute "TYPE", "SUCCESS"

    If m_objContext Is Nothing Then
        Set objIApplManagerDO = New ApplicationManagerDO
        Set objAuditBO = New AuditBO
        Set objIAuditDO = New AuditDO
    Else
        Set objIApplManagerDO = m_objContext.CreateInstance(App.Title & ".ApplicationManagerDO")
        Set objAuditBO = m_objContext.CreateInstance(gstrAUDIT_COMPONENT & ".AuditBO")
        Set objIAuditDO = m_objContext.CreateInstance(gstrAUDIT_COMPONENT & ".AuditDO")
    End If

    'Delegate the call down to the DO.
    objIApplManagerDO.LockCustomersForApplication vxmlRequest
    
    'SR 01/05/2006 : MAR1631 - create AccessAudit record on accessing the application
    Set xmlRequestElem = m_objXmlAssist.GetRequestNode(vxmlRequest.ownerDocument)
    Set xmlAccessAudit = xmlOut.createElement("ACCESSAUDIT")
    xmlRequestElem.appendChild xmlAccessAudit
    xmlAccessAudit.appendChild vxmlRequest.selectSingleNode("//APPLICATIONNUMBER").cloneNode(True)
        
    Set xmlTempElem = xmlOut.createElement("AUDITRECORDTYPE")
    xmlTempElem.Text = objIAuditDO.GetApplicationLockValueId()
    xmlAccessAudit.appendChild xmlTempElem
    
    Set xmlTempElem = xmlOut.createElement("SUCCESSINDICATOR")
    xmlTempElem.Text = "1"
    xmlAccessAudit.appendChild xmlTempElem
    
    strResponse = objAuditBO.CreateAccessAudit(xmlRequestElem.xml)
    m_objErrAssist.CheckXMLResponse strResponse, True
    'SR 01/05/2006 : MAR1631 - End

    Set IApplicationManagerTxBO_LockCustomersForApplication = xmlResponseElem

    If Not m_objContext Is Nothing Then
        m_objContext.SetComplete
    End If

LockCustomersForApplicationExit:

    Set xmlOut = Nothing
    Set xmlResponseElem = Nothing
    Set xmlTempResponseNode = Nothing
    Set objIApplManagerDO = Nothing
    Set objAuditBO = Nothing
    Set objIAuditDO = Nothing
    Set xmlRequestElem = Nothing
    Set xmlAccessAudit = Nothing
    Set xmlTempElem = Nothing
    
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

    Set IApplicationManagerTxBO_LockCustomersForApplication = m_objErrAssist.CreateErrorResponseEx()

    If Not m_objContext Is Nothing Then
        m_objContext.SetAbort
    End If

    ' go to clean-up section
    Resume LockCustomersForApplicationExit

End Function

Private Function IApplicationManagerTxBO_MaintainCustomerRoles( _
                    ByVal vxmlRequest As IXMLDOMElement) As IXMLDOMNode
' header ----------------------------------------------------------------------------------
' description:
' pass:
'   vstrXMLRequest
' return:
'   n/a
' Raise Errors:
'------------------------------------------------------------------------------------------
On Error GoTo MaintainCustomerRolesVbErr

    Const strFunctionName As String = "IApplicationManagerTxBO_MaintainCustomerRoles"

    Dim xmlElement As IXMLDOMElement
    Dim xmlDoc As FreeThreadedDOMDocument40
    Dim xmlOut As New FreeThreadedDOMDocument40
    xmlOut.validateOnParse = False
    xmlOut.setProperty "NewParser", True
    Dim xmlResponseElem As IXMLDOMElement
    Dim xmlTempResponseNode As IXMLDOMNode
    
    Dim objIApplicationBO As IApplicationBO
    Dim objApplBO As ApplicationBO
    Dim objGlobalParameterBO As GlobalParameterBO
    Dim objComboDO As ComboDO 'BMIDS01026
    
    Set xmlResponseElem = xmlOut.createElement("RESPONSE")
    xmlOut.appendChild xmlResponseElem
    xmlResponseElem.setAttribute "TYPE", "SUCCESS"

    '------------------------------------------------------------------------------------------
    ' check application is not a further advance
    '------------------------------------------------------------------------------------------
    ' APS UNIT TEST REF 1 - When validating for a further advance only do this
    ' when the legacy-system is available

    If m_objContext Is Nothing Then
        Set objGlobalParameterBO = New GlobalParameterBO
        Set objIApplicationBO = New ApplicationBO
        Set objApplBO = New ApplicationBO
        'BMIDS01026
        Set objComboDO = New ComboDO
    Else
        Set objGlobalParameterBO = _
            m_objContext.CreateInstance(gstrBASE_COMPONENT & ".GlobalParameterBO")
        Set objIApplicationBO = _
            m_objContext.CreateInstance(cstrCOMPONENT_NAME & ".ApplicationBO")
        Set objApplBO = _
            m_objContext.CreateInstance(cstrCOMPONENT_NAME & ".ApplicationBO")
        'BMIDS01026
        Set objComboDO = m_objContext.CreateInstance(gstrBASE_COMPONENT & ".ComboDO")
    End If

    Dim strLegacyCustomer As String
    Dim strXmlResponse As String
 
    strXmlResponse = objGlobalParameterBO.GetCurrentParameter("FindLegacyCustomer")
    Set xmlDoc = m_objXmlAssist.load(strXmlResponse, TypeName(Me), strFunctionName)
    strLegacyCustomer = m_objXmlAssist.GetTagValue(xmlDoc.documentElement, "BOOLEAN")
    
    ' Indicates whether the interface to the organisation's legacy systems is available
    If strLegacyCustomer = "1" Then

        Dim strTypeOfApp As String
        strTypeOfApp = m_objXmlAssist.GetTagValue(vxmlRequest, "TYPEOFAPPLICATION")

        'BMIDS01026 Throw an error if TypeOfMortgage has validation types F and M, but not S or P
        'If (objApplBO.IsFurtherAdvance(strTypeOfApp)) Then
        If (objComboDO.IsItemInValidation("TypeOfMortgage", strTypeOfApp, "F") And objComboDO.IsItemInValidation("TypeOfMortgage", strTypeOfApp, "M")) Then
            If Not (objComboDO.IsItemInValidation("TypeOfMortgage", strTypeOfApp, "S") Or objComboDO.IsItemInValidation("TypeOfMortgage", strTypeOfApp, "P")) Then
        'BMIDS01026 End
                '"A customer may not be reordered within, assigned to, or removed from this type of mortgage application"
                m_objErrAssist.ThrowError TypeName(Me), strFunctionName, 168
            End If
        End If 'BMIDS01026
    End If

    '---------------------------------------------------------------------------------
    ' update ApplicationFactFind applicant counts
    '---------------------------------------------------------------------------------
    Set xmlTempResponseNode = objIApplicationBO.UpdateApplicantCounts(vxmlRequest)
    m_objErrAssist.CheckResponse xmlTempResponseNode, xmlResponseElem, True

    '------------------------------------------------------------------------------------------
    ' update application correspondence salutation
    '------------------------------------------------------------------------------------------
    Set xmlTempResponseNode = objIApplicationBO.UpdateCorrespondenceSalutation( _
                                                                vxmlRequest)
    m_objErrAssist.CheckResponse xmlTempResponseNode, xmlResponseElem, True
    
    Set IApplicationManagerTxBO_MaintainCustomerRoles = xmlResponseElem
    
    If Not m_objContext Is Nothing Then
        m_objContext.SetComplete
    End If
    
MaintainCustomerRolesExit:

    Set objGlobalParameterBO = Nothing
    Set objIApplicationBO = Nothing
    Set objApplBO = Nothing
    
    Set xmlElement = Nothing
    Set xmlDoc = Nothing
    Set xmlOut = Nothing
    Set xmlResponseElem = Nothing
    Set xmlTempResponseNode = Nothing
    Set objComboDO = Nothing 'BMIDS01026
    
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

    Set IApplicationManagerTxBO_MaintainCustomerRoles = _
                                        m_objErrAssist.CreateErrorResponseEx()

    If Not m_objContext Is Nothing Then
        m_objContext.SetAbort
    End If
        
    ' go to clean-up section
    Resume MaintainCustomerRolesExit
    
End Function

Private Function IApplicationManagerTxBO_SaveCustomerAndEmploymentDetails( _
                    ByVal vxmlRequest As IXMLDOMElement) As IXMLDOMNode

On Error GoTo SaveVbErr

    Const strFunctionName As String = "IApplicationManagerTxBO_SaveCustomerAndEmploymentDetails"

    Dim xmlOut As New FreeThreadedDOMDocument40
    xmlOut.validateOnParse = False
    xmlOut.setProperty "NewParser", True
    Dim xmlResponseElem As IXMLDOMElement, xmlTempResponseNode As IXMLDOMNode
    
    Set xmlResponseElem = xmlOut.createElement("RESPONSE")
    xmlOut.appendChild xmlResponseElem
    xmlResponseElem.setAttribute "TYPE", "SUCCESS"
    
    Dim objICustClassDef As IomCustClassDef
    Set objICustClassDef = New omCustClassDef
    
    Dim booCustOperation As BOOPERATIONTYPE
    booCustOperation = m_objXmlAssist.DetermineOperationEx(vxmlRequest, objICustClassDef.LoadCustomerVersionData)
    
    Select Case booCustOperation
        Case booCreate
            Set xmlTempResponseNode = CreateCustomerAndEmploymentDetails(vxmlRequest)
            m_objErrAssist.CheckResponse xmlTempResponseNode, xmlResponseElem, True
            xmlResponseElem.appendChild xmlTempResponseNode.selectSingleNode(".//CUSTOMERNUMBER")
        Case booUpdate
            Set xmlTempResponseNode = UpdateCustomerAndEmploymentDetails(vxmlRequest)
            m_objErrAssist.CheckResponse xmlTempResponseNode, xmlResponseElem, True
        Case Else
            ' N/A
    End Select
        
    Set IApplicationManagerTxBO_SaveCustomerAndEmploymentDetails = xmlResponseElem
    
    If Not m_objContext Is Nothing Then
        m_objContext.SetComplete
    End If
    
SaveVbExit:
    Set xmlOut = Nothing
    Set xmlResponseElem = Nothing
    Set xmlTempResponseNode = Nothing
    Set objICustClassDef = Nothing
    
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

    Set IApplicationManagerTxBO_SaveCustomerAndEmploymentDetails = _
                                    m_objErrAssist.CreateErrorResponseEx()

    If Not m_objContext Is Nothing Then
        m_objContext.SetAbort
    End If
    
    ' go to clean-up section
    Resume SaveVbExit

End Function

Private Function CreateCustomerAndEmploymentDetails( _
                    ByVal vxmlRequest As IXMLDOMNode) As IXMLDOMNode
    
On Error GoTo CreateVbErr

    Const strFunctionName As String = "CreateCustomerAndEmploymentDetails"
       
    Dim xmlOut As New FreeThreadedDOMDocument40, xmlDoc As FreeThreadedDOMDocument40
    xmlOut.validateOnParse = False
    xmlOut.setProperty "NewParser", True
    Dim xmlResponseElem As IXMLDOMElement, xmlTempResponseNode As IXMLDOMNode
    Dim xmlTableNode As IXMLDOMNode, xmlNode As IXMLDOMNode
    Dim xmlNewRequest As IXMLDOMNode
    
    Dim objCustomerBO As CustomerBO, objCEBO As CustomerEmploymentBO
    Dim objIApplicationBO As IApplicationBO
    Dim objIApplMgrBO As IApplicationManagerBO
    Set objIApplMgrBO = New ApplicationManagerBO
    
    Set xmlResponseElem = xmlOut.createElement("RESPONSE")
    xmlOut.appendChild xmlResponseElem
    xmlResponseElem.setAttribute "TYPE", "SUCCESS"
    
    If m_objContext Is Nothing Then
        Set objCustomerBO = New CustomerBO
        Set objCEBO = New CustomerEmploymentBO
        Set objIApplicationBO = New ApplicationBO
    Else
        Set objCustomerBO = m_objContext.CreateInstance(gstrCUSTOMER_COMPONENT & ".CustomerBO")
        Set objCEBO = m_objContext.CreateInstance(gstrCUSTOMER_EMPLOYMENT & ".CustomerEmploymentBO")
        Set objIApplicationBO = m_objContext.CreateInstance(App.Title & ".ApplicationBO")
    End If

    Dim strResponse As String, strCustomerNumber As String
    
    ' Create Customer data and store the CustomerNumber generated
    
    strResponse = objCustomerBO.CreateCustomerDetails(vxmlRequest.xml)
    Set xmlDoc = m_objXmlAssist.load(strResponse, TypeName(Me), strFunctionName)
    m_objErrAssist.CheckResponse xmlDoc.documentElement, xmlResponseElem, True
    
    strCustomerNumber = xmlDoc.selectSingleNode(".//CUSTOMERNUMBER").Text
    
    ' Create link between customer application
    
    Set xmlTableNode = vxmlRequest.selectSingleNode(".//CUSTOMERROLE")
    If Not xmlTableNode Is Nothing Then
        ' Add customernumber and customerversion number to CustomerRole node
        Set xmlNode = xmlTableNode.selectSingleNode("./CUSTOMERNUMBER")
        If xmlNode Is Nothing Then
            Set xmlNode = xmlDoc.createElement("CUSTOMERNUMBER")
            xmlTableNode.appendChild xmlNode
        End If
        xmlNode.Text = strCustomerNumber
        
        Set xmlNode = xmlTableNode.selectSingleNode("./CUSTOMERVERSIONNUMBER")
        If xmlNode Is Nothing Then
            Set xmlNode = xmlDoc.createElement("CUSTOMERVERSIONNUMBER")
            xmlTableNode.appendChild xmlNode
        End If
        xmlNode.Text = "1"
        
        Set xmlNewRequest = m_objXmlAssist.GetRequestNodeEx(vxmlRequest)
        xmlNewRequest.appendChild xmlTableNode.cloneNode(True)
                
        Set xmlTempResponseNode = IApplicationManagerTxBO_AddCustomerToApplication(xmlNewRequest)
        m_objErrAssist.CheckResponse xmlTempResponseNode, xmlResponseElem, True
    End If
    
    ' Save Customer Relationship
    Set xmlTableNode = vxmlRequest.selectSingleNode(".//CUSTOMERRELATIONSHIP")
    If Not xmlTableNode Is Nothing Then
        ' Add customernumber and customerversion number to CustomerRole node
        Set xmlNode = xmlTableNode.selectSingleNode("./RELTOCUSTOMERNUMBER")
        If xmlNode Is Nothing Then
            Set xmlNode = xmlDoc.createElement("RELTOCUSTOMERNUMBER")
            xmlTableNode.appendChild xmlNode
        End If
        xmlNode.Text = strCustomerNumber
        
        Set xmlNode = xmlTableNode.selectSingleNode("./RELTOCUSTOMERVERSIONNUMBER")
        If xmlNode Is Nothing Then
            Set xmlNode = xmlDoc.createElement("RELTOCUSTOMERVERSIONNUMBER")
            xmlTableNode.appendChild xmlNode
        End If
        xmlNode.Text = "1"
            
        Set xmlNewRequest = m_objXmlAssist.GetRequestNodeEx(vxmlRequest)
        xmlNewRequest.appendChild xmlTableNode.cloneNode(True)
            
        Set xmlTempResponseNode = objIApplicationBO.SaveCustomerRelationship(xmlNewRequest)
        m_objErrAssist.CheckResponse xmlTempResponseNode, xmlResponseElem, True
    End If
    
    ' Save Employment And Employed Details
    Set xmlTableNode = vxmlRequest.selectSingleNode(".//EMPLOYMENT")
    If Not xmlTableNode Is Nothing Then
        Set xmlNode = xmlTableNode.selectSingleNode("./CUSTOMERNUMBER")
        If xmlNode Is Nothing Then
            Set xmlNode = xmlDoc.createElement("CUSTOMERNUMBER")
            xmlTableNode.appendChild xmlNode
        End If
        xmlNode.Text = strCustomerNumber
        
        Set xmlNode = xmlTableNode.selectSingleNode("./CUSTOMERVERSIONNUMBER")
        If xmlNode Is Nothing Then
            Set xmlNode = xmlDoc.createElement("CUSTOMERVERSIONNUMBER")
            xmlTableNode.appendChild xmlNode
        End If
        xmlNode.Text = "1"
            
        Set xmlNewRequest = m_objXmlAssist.GetRequestNodeEx(vxmlRequest)
        xmlNewRequest.appendChild xmlTableNode.cloneNode(True)
            
        ' Set values for Customernumber and CustomerVersionNumber under EmployedDetails
        Set xmlTableNode = xmlNewRequest.selectSingleNode(".//EMPLOYEDDETAILS")
        If Not xmlTableNode Is Nothing Then
            Set xmlNode = xmlTableNode.selectSingleNode(".//CUSTOMERNUMBER")
            If xmlNode Is Nothing Then
                Set xmlNode = xmlDoc.createElement("CUSTOMERNUMBER")
                xmlTableNode.appendChild xmlNode
            End If
            xmlNode.Text = strCustomerNumber
        
            Set xmlNode = xmlTableNode.selectSingleNode(".//CUSTOMERVERSIONNUMBER")
            If xmlNode Is Nothing Then
                Set xmlNode = xmlDoc.createElement("CUSTOMERVERSIONNUMBER")
                xmlTableNode.appendChild xmlNode
            End If
            xmlNode.Text = "1"
        End If
            
        strResponse = objCEBO.SaveEmploymentDetails(xmlNewRequest.xml)
        Set xmlDoc = m_objXmlAssist.load(strResponse, TypeName(Me), strFunctionName)
        m_objErrAssist.CheckResponse xmlDoc.documentElement, xmlResponseElem, True
    End If
    
    ' Add CustomerNumber generated to the Response
    Set xmlNode = xmlOut.createElement("CUSTOMERNUMBER")
    xmlNode.Text = strCustomerNumber
    xmlResponseElem.appendChild xmlNode
    
    Set CreateCustomerAndEmploymentDetails = xmlResponseElem
    
    If Not m_objContext Is Nothing Then
        m_objContext.SetComplete
    End If
    
CreateVbExit:
    Set xmlOut = Nothing
    Set xmlDoc = Nothing
    Set xmlResponseElem = Nothing
    Set xmlTempResponseNode = Nothing
    Set xmlTableNode = Nothing
    Set xmlNode = Nothing
    Set xmlNewRequest = Nothing
    
    Set objCustomerBO = Nothing
    Set objCEBO = Nothing
    Set objIApplMgrBO = Nothing
    
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

    Set CreateCustomerAndEmploymentDetails = m_objErrAssist.CreateErrorResponseEx()

    If Not m_objContext Is Nothing Then
        m_objContext.SetAbort
    End If
    
    ' go to clean-up section
    Resume CreateVbExit

End Function

Private Function UpdateCustomerAndEmploymentDetails( _
                ByVal vxmlRequest As IXMLDOMNode) As IXMLDOMNode

On Error GoTo UpdateVbErr

    Const strFunctionName As String = "UpdateCustomerAndEmploymentDetails"
       
    Dim xmlOut As New FreeThreadedDOMDocument40, xmlDoc As FreeThreadedDOMDocument40
    xmlOut.validateOnParse = False
    xmlOut.setProperty "NewParser", True
    Dim xmlResponseElem As IXMLDOMElement, xmlTempResponseNode As IXMLDOMNode
    Dim xmlNewRequest As IXMLDOMNode, xmlTableNode As IXMLDOMNode
    Dim xmlNode As IXMLDOMNode
            
    Set xmlResponseElem = xmlOut.createElement("RESPONSE")
    xmlOut.appendChild xmlResponseElem
    xmlResponseElem.setAttribute "TYPE", "SUCCESS"
    
    Dim objCustomerBO As CustomerBO, objCEBO As CustomerEmploymentBO
    Dim objIApplicationBO As IApplicationBO
        
    If m_objContext Is Nothing Then
        Set objCustomerBO = New CustomerBO
        Set objCEBO = New CustomerEmploymentBO
        Set objIApplicationBO = New ApplicationBO
    Else
        Set objCustomerBO = m_objContext.CreateInstance(gstrCUSTOMER_COMPONENT & ".CustomerBO")
        Set objCEBO = m_objContext.CreateInstance(gstrCUSTOMER_EMPLOYMENT & ".CustomerEmploymentBO")
        Set objIApplicationBO = m_objContext.CreateInstance(gstrAPPLICATION_COMPONENT & ".ApplicationBO")
    End If
        
    Dim strResponse As String
    
    ' Update Personal Details
    strResponse = objCustomerBO.UpdatePersonalDetails(vxmlRequest.xml)
    Set xmlDoc = m_objXmlAssist.load(strResponse, TypeName(Me), strFunctionName)
    m_objErrAssist.CheckResponse xmlDoc.documentElement, xmlResponseElem, True
    
    ' Update Correspondence Salutation
    Set xmlNewRequest = m_objXmlAssist.GetRequestNodeEx(vxmlRequest)
    Set xmlTableNode = xmlDoc.createElement("APPLICATION")
    xmlNewRequest.appendChild xmlTableNode
        
    xmlTableNode.appendChild vxmlRequest.selectSingleNode(".//APPLICATIONNUMBER").cloneNode(True)
    xmlTableNode.appendChild vxmlRequest.selectSingleNode(".//APPLICATIONFACTFINDNUMBER").cloneNode(True)

    Set xmlTempResponseNode = objIApplicationBO.UpdateCorrespondenceSalutation(xmlNewRequest)
    m_objErrAssist.CheckResponse xmlTempResponseNode, xmlResponseElem, True
    
    ' Update Employment and Employed details
    Set xmlTableNode = vxmlRequest.selectSingleNode(".//EMPLOYMENT")
    If Not xmlTableNode Is Nothing Then
        Set xmlNewRequest = m_objXmlAssist.GetRequestNodeEx(vxmlRequest)
        xmlNewRequest.appendChild xmlTableNode.cloneNode(True)

        strResponse = objCEBO.SaveEmploymentDetails(xmlNewRequest.xml)
        Set xmlDoc = m_objXmlAssist.load(strResponse, TypeName(Me), strFunctionName)
        m_objErrAssist.CheckResponse xmlDoc.documentElement, xmlResponseElem, True
    End If
    
    ' Save Customer relationship
    Set xmlTableNode = vxmlRequest.selectSingleNode(".//CUSTOMERRELATIONSHIP")
    If Not xmlTableNode Is Nothing Then
        Set xmlTempResponseNode = objIApplicationBO.SaveCustomerRelationship(vxmlRequest)
        m_objErrAssist.CheckResponse xmlTempResponseNode, xmlResponseElem, True
    End If
    
    ' Delete Other names(if any), if new value for 'OtherNamesIndicator' is 0
    Dim strAliasIndicator As String
    
    strAliasIndicator = m_objXmlAssist.GetTagValue(vxmlRequest, "ALIASINDICATOR")
    
    If strAliasIndicator = "0" Then
        Set xmlNewRequest = m_objXmlAssist.GetRequestNodeEx(vxmlRequest)
        Set xmlTableNode = xmlDoc.createElement("ALIAS")
        xmlNewRequest.appendChild xmlTableNode
        
        xmlTableNode.appendChild vxmlRequest.selectSingleNode(".//CUSTOMERNUMBER").cloneNode(True)
        xmlTableNode.appendChild vxmlRequest.selectSingleNode(".//CUSTOMERVERSIONNUMBER").cloneNode(True)
        
        strResponse = objCustomerBO.FindAliasList(xmlNewRequest.xml)
        Set xmlDoc = m_objXmlAssist.load(strResponse, TypeName(Me), strFunctionName)
        
        On Error GoTo RecordNotFoundErr
        m_objErrAssist.CheckResponse xmlDoc.documentElement, xmlResponseElem, True
        On Error GoTo UpdateVbErr
        
        Set xmlNode = xmlDoc.selectSingleNode(".//ALIASPERSONLIST")
        
        If Not xmlNode Is Nothing Then
            Set xmlNewRequest = m_objXmlAssist.GetRequestNodeEx(vxmlRequest)
            xmlNewRequest.appendChild xmlNode
            strResponse = objCustomerBO.DeleteAlias(xmlNewRequest.xml)
            Set xmlDoc = m_objXmlAssist.load(strResponse, TypeName(Me), strFunctionName)
            m_objErrAssist.CheckResponse xmlDoc.documentElement, xmlResponseElem, True
        End If
    End If
    
    Set UpdateCustomerAndEmploymentDetails = xmlResponseElem
    
    If Not m_objContext Is Nothing Then
        m_objContext.SetComplete
    End If

UpdateVbExit:
    Set xmlOut = Nothing
    Set xmlDoc = Nothing
    Set xmlResponseElem = Nothing
    Set xmlTempResponseNode = Nothing
    Set xmlNewRequest = Nothing
    Set xmlTableNode = Nothing
    Set xmlNode = Nothing
            
    Set objCustomerBO = Nothing
    Set objCEBO = Nothing
    Set objIApplicationBO = Nothing
    
    Exit Function
RecordNotFoundErr:
    If m_objErrAssist.GetOmigaErrorNumber(Err.Number) = oeRecordNotFound Then
        Resume Next
    End If
    
UpdateVbErr:
    If m_objErrAssist.IsWarning = True Then
        m_objErrAssist.AddWarning xmlResponseElem
        Resume Next
    End If

    m_objErrAssist.AddToErrSource strFunctionName
    
    If m_objErrAssist.IsSystemError = True Then
        m_objErrAssist.LogError TypeName(Me), strFunctionName, Err.Number, Err.Description
    End If

    Set UpdateCustomerAndEmploymentDetails = m_objErrAssist.CreateErrorResponseEx()

    If Not m_objContext Is Nothing Then
        m_objContext.SetAbort
    End If
    
    ' go to clean-up section
    Resume UpdateVbExit
    
End Function

Private Function IApplicationManagerTxBO_UnlockApplicationAndCustomers( _
                ByVal vxmlRequest As IXMLDOMElement) As IXMLDOMNode
' header -------------------------------------------------------------------------
' description:  Remove the locks on application and related customers
'               invokes appropriate Data Object(s) to perform persistance
' pass:         vxmlRequest  xml Request element
' return:       xml Response Node
'---------------------------------------------------------------------------------
On Error GoTo UnlockApplicationAndCustomersVbErr

    Const strFunctionName As String = "IApplicationManagerTxBO_UnlockApplicationAndCustomers"

    Dim xmlOut As FreeThreadedDOMDocument40
    Dim xmlResponseElem As IXMLDOMElement
    Dim objIApplManagerDO As IApplicationManagerDO
    Dim objAuditBO As AuditBO, objIAuditDO As IAuditDO 'SR MAR1631
    Dim xmlRequestElem As IXMLDOMElement, xmlAccessAudit As IXMLDOMElement, xmlTempElem As IXMLDOMElement 'SR MAR1631
    Dim strResponse As String 'SR MAR1631


    If m_objContext Is Nothing Then
        Set objIApplManagerDO = New ApplicationManagerDO
        Set objAuditBO = New AuditBO
        Set objIAuditDO = New AuditDO
    Else
        Set objIApplManagerDO = m_objContext.CreateInstance(App.Title & ".ApplicationManagerDO")
        Set objAuditBO = m_objContext.CreateInstance(gstrAUDIT_COMPONENT & ".AuditBO")
        Set objIAuditDO = m_objContext.CreateInstance(gstrAUDIT_COMPONENT & ".AuditDO")

    End If
    
    Set xmlOut = New FreeThreadedDOMDocument40
    xmlOut.validateOnParse = False
    xmlOut.setProperty "NewParser", True
    Set xmlResponseElem = xmlOut.createElement("RESPONSE")
    xmlOut.appendChild xmlResponseElem
    xmlResponseElem.setAttribute "TYPE", "SUCCESS"
                    
    'Delegate the call down to the DO.
    objIApplManagerDO.UnlockApplicationAndCustomers vxmlRequest
    
    'SR 01/05/2006 : MAR1631 - create AccessAudit record on after unlocking application
    Set xmlRequestElem = m_objXmlAssist.GetRequestNode(vxmlRequest.ownerDocument)
    Set xmlAccessAudit = xmlOut.createElement("ACCESSAUDIT")
    xmlRequestElem.appendChild xmlAccessAudit
    xmlAccessAudit.appendChild vxmlRequest.selectSingleNode("//APPLICATIONNUMBER").cloneNode(True)
        
    Set xmlTempElem = xmlOut.createElement("AUDITRECORDTYPE")
    xmlTempElem.Text = objIAuditDO.GetApplicationReleaseValueId()
    xmlAccessAudit.appendChild xmlTempElem
    
    Set xmlTempElem = xmlOut.createElement("SUCCESSINDICATOR")
    xmlTempElem.Text = "1"
    xmlAccessAudit.appendChild xmlTempElem
    
    strResponse = objAuditBO.CreateAccessAudit(xmlRequestElem.xml)
    m_objErrAssist.CheckXMLResponse strResponse, True
    'SR 01/05/2006 : MAR1631 - End

    Set IApplicationManagerTxBO_UnlockApplicationAndCustomers = xmlResponseElem
    
    If Not m_objContext Is Nothing Then
        m_objContext.SetComplete
    End If
    
UnlockApplicationAndCustomersExit:
    
    Set objIApplManagerDO = Nothing

    Set xmlOut = Nothing
    Set xmlResponseElem = Nothing
    Set objAuditBO = Nothing
    Set objIAuditDO = Nothing
    Set xmlRequestElem = Nothing
    Set xmlAccessAudit = Nothing
    Set xmlTempElem = Nothing
    
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

    Set IApplicationManagerTxBO_UnlockApplicationAndCustomers = m_objErrAssist.CreateErrorResponseEx()

    If Not m_objContext Is Nothing Then
        m_objContext.SetAbort
    End If
    
    ' go to clean-up section
    Resume UnlockApplicationAndCustomersExit

End Function

Private Function IApplicationManagerTxBO_UpdateCustomerRole(ByVal vxmlRequest As IXMLDOMElement) As IXMLDOMNode
' header -------------------------------------------------------------------------
' description:  Update multiple instances of the persistant data associated with
'               the table 'CustomerRole'
'               invokes appropriate Data Object(s) to perform persistance
' pass:         vxmlRequest  xml Request element
' return:       xml Response Node
'---------------------------------------------------------------------------------
On Error GoTo UpdateVbErr

    Const strFunctionName As String = "IApplicationManagerTxBO_UpdateCustomerRole"

    Dim xmlNodeList As IXMLDOMNodeList
    Dim xmlNode As IXMLDOMNode
    Dim xmlElement As IXMLDOMElement
    Dim xmlOut As New FreeThreadedDOMDocument40
    xmlOut.validateOnParse = False
    xmlOut.setProperty "NewParser", True
    Dim xmlResponseElem As IXMLDOMElement

    Dim objIApplManagerDO As IApplicationManagerDO
    
    Set xmlResponseElem = xmlOut.createElement("RESPONSE")
    xmlOut.appendChild xmlResponseElem
    xmlResponseElem.setAttribute "TYPE", "SUCCESS"
                
    ' Check if the element passed in is the one with the table name else
    ' look below the node that is passed in
    If vxmlRequest.nodeName = cstTABLE_CUSTOMERROLE Then
        Set xmlElement = vxmlRequest
    Else
        Set xmlNodeList = vxmlRequest.selectNodes("CUSTOMERROLELIST/" & cstTABLE_CUSTOMERROLE)  'EP2_56 GHun
    End If
    
    If xmlElement Is Nothing Then
        If xmlNodeList.length = 0 Then
            m_objErrAssist.ThrowError TypeName(Me), _
                                  strFunctionName, _
                                  oeMissingPrimaryTag, _
                                  cstTABLE_CUSTOMERROLE & " tag not found"
        End If
    End If
   
    If m_objContext Is Nothing Then
        Set objIApplManagerDO = New ApplicationManagerDO
    Else
        Set objIApplManagerDO = m_objContext.CreateInstance(App.Title & ".ApplicationManagerDO")
    End If
    
    If Not xmlElement Is Nothing Then
        objIApplManagerDO.UpdateCustomerRole xmlElement
    Else
        For Each xmlNode In xmlNodeList
            objIApplManagerDO.UpdateCustomerRole xmlNode
        Next xmlNode
    End If
    
    'EP2_56 GHun
    For Each xmlElement In vxmlRequest.selectNodes("REMOVEDTOECUSTOMERLIST/REMOVEDTOECUSTOMER")
        objIApplManagerDO.CreateRemovedToECustomer xmlElement
    Next
    'EP2_56 End
    
    Set IApplicationManagerTxBO_UpdateCustomerRole = xmlResponseElem
    
    If Not m_objContext Is Nothing Then
        m_objContext.SetComplete
    End If
    
UpdateExit:
    
    Set objIApplManagerDO = Nothing
    Set xmlNodeList = Nothing
    Set xmlElement = Nothing
    Set xmlOut = Nothing
    Set xmlResponseElem = Nothing
    
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

    Set IApplicationManagerTxBO_UpdateCustomerRole = m_objErrAssist.CreateErrorResponseEx()

    If Not m_objContext Is Nothing Then
        m_objContext.SetAbort
    End If
    
    ' go to clean-up section
    Resume UpdateExit

End Function

Private Function IApplicationManagerTxBO_UpdateCustomerRoles( _
                ByVal vxmlRequest As IXMLDOMElement) As IXMLDOMNode
' header ----------------------------------------------------------------------------------
' description:
' pass:
'   vstrXMLRequest
' return:
'   n/a
' Raise Errors:
'------------------------------------------------------------------------------------------
On Error GoTo UpdateCustomerRolesVbErr

    Const strFunctionName As String = "IApplicationManagerTxBO_UpdateCustomerRoles"

    Dim xmlDoc As New FreeThreadedDOMDocument40
    xmlDoc.validateOnParse = False
    xmlDoc.setProperty "NewParser", True
    Dim xmlOut As New FreeThreadedDOMDocument40
    xmlOut.validateOnParse = False
    xmlOut.setProperty "NewParser", True
    Dim xmlResponseElem As IXMLDOMElement
    Dim xmlTempResponseNode As IXMLDOMNode
    Dim xmlInputElem As IXMLDOMElement  'Input to MaintainCustomerRoles
    Dim xmlApplElem As IXMLDOMElement
    Dim xmlElement As IXMLDOMElement
    
    Dim objIApplManagerBO As IApplicationManagerBO
            
    Set xmlResponseElem = xmlOut.createElement("RESPONSE")
    xmlOut.appendChild xmlResponseElem
    xmlResponseElem.setAttribute "TYPE", "SUCCESS"
    
    '-------------------------------------------------------------------------
    ' update the CustomerRole table
    '-------------------------------------------------------------------------
    If m_objContext Is Nothing Then
        Set objIApplManagerBO = New ApplicationManagerBO
    Else
        Set objIApplManagerBO = _
            m_objContext.CreateInstance(cstrCOMPONENT_NAME & ".ApplicationManagerBO")
    End If
    
    Set xmlTempResponseNode = objIApplManagerBO.UpdateCustomerRole(vxmlRequest)
    m_objErrAssist.CheckResponse xmlTempResponseNode, xmlResponseElem, True
    
    '---------------------------------------------------------------------------
    ' validate the change
    '---------------------------------------------------------------------------
    'Build the node to be passed to the method 'MaintainCustomerRoles'
'    Set xmlDoc = m_objXmlAssist.load(vxmlRequest, TypeName(Me), strFunctionName)
    Set xmlInputElem = xmlDoc.createElement("REQUEST")
    Set xmlApplElem = xmlDoc.createElement("APPLICATION")
    xmlInputElem.appendChild xmlApplElem
    
    Dim strTag As String
    strTag = "APPLICATIONNUMBER"
    Set xmlElement = xmlDoc.createElement(strTag)
    xmlElement.Text = vxmlRequest.getElementsByTagName(strTag).Item(0).Text
    xmlApplElem.appendChild xmlElement
    
    strTag = "APPLICATIONFACTFINDNUMBER"
    Set xmlElement = xmlDoc.createElement(strTag)
    xmlElement.Text = vxmlRequest.getElementsByTagName(strTag).Item(0).Text
    xmlApplElem.appendChild xmlElement
    
    strTag = "TYPEOFAPPLICATION"
    Set xmlElement = xmlDoc.createElement(strTag)
    xmlElement.Text = vxmlRequest.getElementsByTagName(strTag).Item(0).Text
    xmlApplElem.appendChild xmlElement
    
    Set xmlTempResponseNode = _
                IApplicationManagerTxBO_MaintainCustomerRoles(xmlInputElem)
    m_objErrAssist.CheckResponse xmlTempResponseNode, xmlResponseElem, True
    
    Set IApplicationManagerTxBO_UpdateCustomerRoles = xmlResponseElem
    
    If Not m_objContext Is Nothing Then
        m_objContext.SetComplete
    End If
    
UpdateCustomerRolesExit:
    Set xmlDoc = Nothing
    Set xmlOut = Nothing
    Set xmlResponseElem = Nothing
    Set xmlTempResponseNode = Nothing
    Set xmlInputElem = Nothing
    Set xmlApplElem = Nothing
    Set xmlElement = Nothing
    Set objIApplManagerBO = Nothing
    
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

    Set IApplicationManagerTxBO_UpdateCustomerRoles = _
                                        m_objErrAssist.CreateErrorResponseEx()

    If Not m_objContext Is Nothing Then
        m_objContext.SetAbort
    End If
        
    ' go to clean-up section
    Resume UpdateCustomerRolesExit

End Function

Private Function IApplicationManagerTxBO_ImportAccountsIntoApplication(ByVal vxmlRequest As IXMLDOMElement) As IXMLDOMNode
    
    On Error GoTo ImportAccountsIntoApplicationVbErr

    Const cstrFunctionName As String = "IApplicationManagerTxBO_ImportAccountsIntoApplication"
    Const cstrTABLE_CUSTOMERLIST As String = "CUSTOMERLIST"

    Dim xmlElement As IXMLDOMElement
    Dim xmlOut As FreeThreadedDOMDocument40
    Dim xmlResponseElem As IXMLDOMElement
    Dim xmlAppNo As IXMLDOMElement
    Dim xmlAppFFNo As IXMLDOMElement
    Dim xmlTempRequestDoc As FreeThreadedDOMDocument40
    Dim xmlRequestElem As IXMLDOMElement
    Dim xmlTableElem As IXMLDOMElement
    Dim xmlTempResponseDoc As FreeThreadedDOMDocument40
    Dim xmlAccount As IXMLDOMElement
    Dim xmlMortgageAccountDoc As FreeThreadedDOMDocument40
    Dim xmlTempElement As IXMLDOMElement
    Dim xmlAccountList As IXMLDOMNodeList
    Dim xmlAccountGuid As IXMLDOMElement
    Dim xmlOtherSystemAccountNumber As IXMLDOMElement
    Dim xmlElemBasedAdmin As IXMLDOMNode
    Dim xmlTempRequestNode As IXMLDOMNode
    Dim xmlPropertyNode As IXMLDOMNode
    Dim xmlAFFElem As IXMLDOMElement
    Dim xmlNewProperty As IXMLDOMNode
    Dim xmlExistingFS As IXMLDOMElement
    Dim xmlSpecialFeaturesNodes   As IXMLDOMNodeList
    Dim xmlSpecialFeature  As IXMLDOMElement
    Dim xmlTempNode  As IXMLDOMNode
        
    Dim objIApplicationBO As IApplicationBO
    Dim objCustFinancialBO As CustomerFinancialBO
    Dim objNewPropertyBO As INewPropertyBO
    Dim objIApplicationDO As IApplicationDO  'SR 13/08/2004 : BMIDS815
    Dim objApplicationTPBO As ApplicationThirdPartyBO  'SR EP2_858
    
    Dim strResponse As String
    Dim strAccountNo As String
    Dim strTypeOfApplication As String  'SR EP2_858
    Dim strCurrAccountNo As String
    Dim strPropertyAddressGuid As String
    Dim strAccountGUID As String
    
    Dim blnFurtherAdvance As Boolean
    Dim blnOtherCustomers As Boolean
    Dim blnMortgageIndicatorSet As Boolean
    Dim blnArrearsIndicatorSet  As Boolean
    Dim blnCreateFinancialSummary As Boolean
        
    'BM0387
    Dim strArrearsHistoryData   As String
    Dim xmlArrearsHistoryList   As IXMLDOMNodeList
    Dim xmlAccountArrears       As IXMLDOMElement
    Dim blnHasArrearsHistory    As Boolean
    Dim iCounter                As Integer
    'BM0387 End
    
    ' PSC 30/01/2007 EP2_1114 - Start
    Dim nDashPos As Integer
    Dim strSortCode As String
    Dim lngErrorNumber As Long
    ' PSC 30/01/2007 EP2_1114 - End
    
    Dim blnReturningCustomer As Boolean     ' PSC 30/01/2007 EP2_1146
    Dim blnNewPortedApplication As Boolean  ' MAH 13/03/2007 EP2_1597
    
    Dim objIAppManagerBO As IApplicationManagerBO
        
    'Initialise Response
    Set xmlOut = New FreeThreadedDOMDocument40
    xmlOut.validateOnParse = False
    xmlOut.setProperty "NewParser", True
    xmlOut.async = False
    Set xmlResponseElem = xmlOut.createElement("RESPONSE")
    xmlOut.appendChild xmlResponseElem
    xmlResponseElem.setAttribute "TYPE", "SUCCESS"
                
    ' Check if the element passed in is the one with the table name else
    ' look below the node that is passed in
    If vxmlRequest.nodeName = cstrTABLE_CUSTOMERLIST Then
        Set xmlElement = vxmlRequest
    Else
        Set xmlElement = vxmlRequest.getElementsByTagName(cstrTABLE_CUSTOMERLIST).Item(0)
    End If
    
    If xmlElement Is Nothing Then
        m_objErrAssist.ThrowError TypeName(Me), _
                                  cstrFunctionName, _
                                  oeMissingPrimaryTag, _
                                  cstrTABLE_CUSTOMERLIST & " tag not found"
    End If
    
    strTypeOfApplication = m_objXmlAssist.GetNodeValue(vxmlRequest, ".//TYPEOFAPPLICATION", True) 'SR Ep2_858
    strAccountNo = m_objXmlAssist.GetNodeValue(vxmlRequest, ".//ACCOUNTNUMBER", True, False)
    
    ' PSC 30/01/2007 EP2_1146
    blnReturningCustomer = (IsValidationType("TypeOfMortgage", strTypeOfApplication, "F") And _
                            IsValidationType("TypeOfMortgage", strTypeOfApplication, "M")) Or _
                            IsValidationType("TypeOfMortgage", strTypeOfApplication, "PSW") Or _
                            IsValidationType("TypeOfMortgage", strTypeOfApplication, "TOE") Or _
                            IsValidationType("TypeOfMortgage", strTypeOfApplication, "NP")
                            
    blnNewPortedApplication = IsValidationType("TypeOfMortgage", strTypeOfApplication, "NP")  ' MAH 13/03/2007 EP2_1597
    
    Set xmlAppNo = m_objXmlAssist.GetNode(vxmlRequest, ".//APPLICATIONNUMBER", True)
    Set xmlAppFFNo = m_objXmlAssist.GetNode(vxmlRequest, ".//APPLICATIONFACTFINDNUMBER", True)
    
    Set objIApplicationBO = m_objContext.CreateInstance(App.Title & ".ApplicationBO")
    Set objCustFinancialBO = m_objContext.CreateInstance(gstrCUSTOMER_FINANCIAL_COMPONENT & ".CustomerFinancialBO")
    Set objIAppManagerBO = m_objContext.CreateInstance(App.Title & ".ApplicationManagerBO")
    Set objIApplicationDO = m_objContext.CreateInstance(App.Title & ".ApplicationDO")  'SR 13/08/2004 - BMIDS815
    Set objApplicationTPBO = m_objContext.CreateInstance(App.Title & ".ApplicationThirdPartyBO")  'SR EP2_858
    
    Set xmlTempRequestDoc = New FreeThreadedDOMDocument40
    xmlTempRequestDoc.validateOnParse = False
    xmlTempRequestDoc.setProperty "NewParser", True
    xmlTempRequestDoc.async = False
    
    
    'BMIDS918 GHun The GetAccountDetails call is now done within omApplicationManagerBO,
    'outside of the transaction
    blnFurtherAdvance = (vxmlRequest.getAttribute("IsFurtherAdvance") = "1")
    
   'SR 16/01/2007 : EP2_858 - removed commented code. Check V2 for removed code.
    
    Set xmlMortgageAccountDoc = New FreeThreadedDOMDocument40
    xmlMortgageAccountDoc.validateOnParse = False
    xmlMortgageAccountDoc.setProperty "NewParser", True
    xmlMortgageAccountDoc.async = False
    xmlMortgageAccountDoc.appendChild vxmlRequest.selectSingleNode("MORTGAGEACCOUNTLIST")
    
   'SR 16/01/2007 : EP2_858 - removed commented code. Check V2 for removed code.
   
    ' Get a list of the accounts from the admin system
    Set xmlElemBasedAdmin = vxmlRequest.selectSingleNode("ADMINSYSTEMRESPONSE")
    Set xmlAccountList = xmlElemBasedAdmin.selectNodes(".//MORTGAGEACCOUNT")
    'BMIDS918 End
    
    '=========================================================================================
    ' Create/Update Financial Summary
    '=========================================================================================
    
    ' Check if there are any arrears in the account details
    'BM0387
    'strMaxBalance = m_objXmlAssist.GetNodeValue(xmlElemBasedAdmin, ".//MAXIMUMBALANCE[. != '']", _
                                                False, False)
    blnHasArrearsHistory = False
    Set xmlArrearsHistoryList = xmlElemBasedAdmin.selectNodes(".//ARREARSHISTORY")
    iCounter = 0
    'Check for any non-blank arrears history records
    Do While (iCounter < xmlArrearsHistoryList.length) And (Not blnHasArrearsHistory)
        
        Set xmlAccountArrears = xmlArrearsHistoryList.Item(iCounter)
        strArrearsHistoryData = xmlGetNodeText(xmlAccountArrears, "MAXIMUMBALANCE") & _
            xmlGetNodeText(xmlAccountArrears, "MAXIMUMNUMBEROFMONTHS") & _
            xmlGetNodeText(xmlAccountArrears, "DATECLEARED") & _
            xmlGetNodeText(xmlAccountArrears, "DESCRIPTIONOFLOAN") & _
            xmlGetNodeText(xmlAccountArrears, "CURRENTYEARSINARREARS") & _
            xmlGetNodeText(xmlAccountArrears, "LASTTWOYEARSINARREARS")
        
        If Len(Trim$(strArrearsHistoryData)) > 0 Then
            blnHasArrearsHistory = True
        End If
        
        iCounter = iCounter + 1
    Loop
    'BM0387 End
    
    ' Update Financial Summary
    Set xmlRequestElem = m_objXmlAssist.GetRequestNodeEx(vxmlRequest)
    Set xmlTempRequestDoc.documentElement = xmlRequestElem
    
    Set xmlTableElem = xmlTempRequestDoc.createElement("FINANCIALSUMMARY")
    xmlRequestElem.appendChild xmlTableElem
    
    xmlTableElem.appendChild xmlAppNo.cloneNode(True)
    xmlTableElem.appendChild xmlAppFFNo.cloneNode(True)
    
    'Determine if a Financial Summary record exists...
    Set xmlExistingFS = objIApplicationBO.GetFinancialSummary(xmlRequestElem)
    m_objErrAssist.CheckResponse xmlExistingFS, xmlResponseElem, True
    
    blnMortgageIndicatorSet = False
    blnArrearsIndicatorSet = False
    
    If xmlExistingFS.getAttribute("TYPE") <> "SUCCESS" Then
        blnCreateFinancialSummary = True
    Else
        If xmlGetNodeAsBoolean(xmlExistingFS, ".//EXISTINGMORTGAGEINDICATOR") Then
            blnMortgageIndicatorSet = True
        End If
        If xmlGetNodeAsBoolean(xmlExistingFS, ".//ARREARSHISTORYINDICATOR") Then
            blnArrearsIndicatorSet = True
        End If
    End If
    
    'Existing Mortgage
    'BM0377 Only update ExistingMortgageIndicator if it is not already set
     If Not blnMortgageIndicatorSet Then
    'BM0377 End
        Set xmlTempElement = xmlTempRequestDoc.createElement("EXISTINGMORTGAGEINDICATOR")
        If xmlAccountList.length > 0 Then
            xmlTempElement.Text = "1"
        Else
            xmlTempElement.Text = "0"
        End If
        xmlTableElem.appendChild xmlTempElement
    End If 'BM0377
        
    'Arrears History
    'BM0377 Only update ArrearsHistoryIndicator if it is not already set
     If Not blnArrearsIndicatorSet Then
    'BM0377 End
        Set xmlTempElement = xmlTempRequestDoc.createElement("ARREARSHISTORYINDICATOR")
        'BM0387
        'If Len(strMaxBalance) > 0 Then
        If blnHasArrearsHistory Then
        'BM0387 End
            xmlTempElement.Text = "1"
        Else
            xmlTempElement.Text = "0"
        End If
        xmlTableElem.appendChild xmlTempElement
    End If 'BM0377
    
    '...and Create or Update accordingly
    If blnCreateFinancialSummary Then
        Set xmlTempElement = objIApplicationBO.CreateFinancialSummary(xmlRequestElem)
    Else
        'BM0377 No need to update FinancialSummary if both indicators are already set
        If Not (blnMortgageIndicatorSet And blnArrearsIndicatorSet) Then
        'BM0377 End
            Set xmlTempElement = objIApplicationBO.UpdateFinancialSummary(xmlRequestElem)
        End If 'BM0377
    End If
    If Not xmlTempElement Is Nothing Then
        m_objErrAssist.CheckResponse xmlTempElement, xmlResponseElem, True
    End If
    
            
    '=========================================================================================
    ' Create required Mortgage Accounts
    '=========================================================================================
    For Each xmlAccount In xmlAccountList
        
        'Create CustomerAddress record -------------------------------------------------------
        Set xmlPropertyNode = m_objXmlAssist.GetNode(xmlAccount, "PROPERTY").cloneNode(True)
        strPropertyAddressGuid = ImportPropertyAddress(vxmlRequest, xmlPropertyNode, xmlResponseElem)
        
        'Create Account and Loans ------------------------------------------------------------
        strAccountGUID = ImportAccountAndLoans(vxmlRequest, _
                                               xmlAccount, _
                                               xmlMortgageAccountDoc, _
                                               xmlResponseElem, _
                                               blnOtherCustomers, _
                                               "", _
                                               strPropertyAddressGuid)
        'Create Arrears History --------------------------------------------------------------
        ImportArrearsHistory vxmlRequest, xmlAccount, strAccountGUID, xmlResponseElem
        
        ' If this is a further advance and this account is the one it is based on then
        ' update the application with the account guid and create the new property and
        ' new property address
        strCurrAccountNo = m_objXmlAssist.GetNodeValue(xmlAccount, ".//ACCOUNTNUMBER")
        
        ' PSC 30/01/2007 EP2_1146
        If blnReturningCustomer = True And strCurrAccountNo = strAccountNo Then
        
            ' Update application with AccountGuid of the 'parent' mortgage account -----------

            Set xmlTableElem = xmlTempRequestDoc.createElement("APPLICATION")
            xmlRequestElem.replaceChild xmlTableElem, xmlRequestElem.firstChild
            xmlTableElem.appendChild xmlAppNo.cloneNode(True)
            xmlTableElem.appendChild xmlAppFFNo.cloneNode(True)
            Set xmlAccountGuid = xmlTempRequestDoc.createElement("ACCOUNTGUID")
            xmlAccountGuid.Text = strAccountGUID
            xmlTableElem.appendChild xmlAccountGuid
            
            ' AQR SYS4740 update the other system account number with the value received from the other system
            Set xmlOtherSystemAccountNumber = xmlTempRequestDoc.createElement("OTHERSYSTEMACCOUNTNUMBER")
            xmlOtherSystemAccountNumber.Text = strAccountNo
            xmlTableElem.appendChild xmlOtherSystemAccountNumber
            
            Set xmlTempElement = objIApplicationBO.Update(xmlRequestElem) 'SR EP2_858 :
            m_objErrAssist.CheckResponse xmlTempElement, xmlResponseElem, True
            
            'SR EP2_858 : 15/01/2007 - Update Application Fact Find
            Set xmlTempRequestNode = vxmlRequest.cloneNode(False)
            Set xmlAFFElem = xmlTempRequestDoc.createElement("APPLICATIONFACTFIND")
            xmlTempRequestNode.appendChild xmlAFFElem
            xmlAFFElem.appendChild xmlAppNo.cloneNode(True)
            xmlAFFElem.appendChild xmlAppFFNo.cloneNode(True)
            
            Set xmlTempElement = xmlTempRequestDoc.createElement("NATUREOFLOAN")
            xmlTempElement.Text = m_objXmlAssist.GetTagValue(xmlAccount, "ORIGINALNATUREOFLOAN")
            xmlAFFElem.appendChild xmlTempElement
            
            Set xmlTempElement = xmlTempRequestDoc.createElement("PRODUCTSCHEME")
            xmlTempElement.Text = m_objXmlAssist.GetTagValue(xmlAccount, "ORIGINALCREDITSCHEME")
            xmlAFFElem.appendChild xmlTempElement
            
            Set xmlTempElement = xmlTempRequestDoc.createElement("APPLICATIONINCOMESTATUS")
            xmlTempElement.Text = m_objXmlAssist.GetElementText(xmlAccount, ".//MORTGAGELOAN/ORIGINALINCOMESTATUS")
            xmlAFFElem.appendChild xmlTempElement
            
            Set xmlTempElement = objIApplicationBO.UpdateApplicationFactFind(xmlTempRequestNode)
            m_objErrAssist.CheckResponse xmlTempElement, xmlResponseElem, True
            'SR EP2_858 : 15/01/2007 - End
          If Not blnNewPortedApplication Then ' MAH 02/04/2007 EP2_1597
            'Save New Property ---------------------------------------------------------------
            xmlTableElem.removeChild xmlAccountGuid
            xmlTableElem.removeChild xmlOtherSystemAccountNumber
            Set objNewPropertyBO = m_objContext.CreateInstance(App.Title & ".NewPropertyBO")
            m_objXmlAssist.ChangeNodeName xmlTableElem, "APPLICATION", "NEWPROPERTY"
            
            'BMIDS00563 Check if NewProperty already exists
            Set xmlNewProperty = objNewPropertyBO.GetNewPropertyGeneral(xmlTempRequestDoc.documentElement)
            
            'Only create a new property record for the application if there isn't one already
            'BMIDS01053 Need to use GetOmigaErrorNumber before checking against constant
            If m_objErrAssist.GetOmigaErrorNumber(m_objErrAssist.CheckXMLResponse(xmlNewProperty.xml)) = oeRecordNotFound Then
            'BMIDS00563 End
          
                'SYS4637 - Append the property details from this account. Copy all the child
                'element-values of the PROPERTY element into the table element.
                For Each xmlTempElement In xmlAccount.selectSingleNode(".//PROPERTY").childNodes
                    'Do not copy any child element structures.
                    If xmlTempElement.hasChildNodes = True Then
                        If xmlTempElement.firstChild.nodeType <> NODE_ELEMENT Then
                            xmlTableElem.appendChild xmlTempElement
                        End If
                    End If
                Next xmlTempElement
                'SYS4637 - End.
                                
                'BM0233 PROPERTYTENURE node must be called TENURETYPE for it to be saved
                Set xmlTempElement = xmlTableElem.selectSingleNode("PROPERTYTENURE") 'SR EP2_858
                If Not xmlTempElement Is Nothing Then 'SR EP2_858
                    xmlChangeNodeName xmlTempElement, "PROPERTYTENURE", "TENURETYPE"
                End If
                'BM0233 End
                      
                Set xmlTempElement = objNewPropertyBO.CreateNewProperty(xmlTempRequestDoc.documentElement)
                m_objErrAssist.CheckResponse xmlTempElement, xmlResponseElem, True
            
            End If  'BMIDS00563 End
            
            'Save New Property Address --------------------------------------------------------
            m_objXmlAssist.ChangeNodeName xmlTableElem, "NEWPROPERTY", "NEWPROPERTYADDRESS"
            
            'BMIDS00563 Check if NewPropertyAddress already exists
            Set xmlNewProperty = objNewPropertyBO.GetNewPropertyAddress(xmlTempRequestDoc.documentElement)
            
            'Only create a new property address record if there isn't one already
            'BMIDS01053 Need to use GetOmigaErrorNumber before checking against constant
            If m_objErrAssist.GetOmigaErrorNumber(m_objErrAssist.CheckXMLResponse(xmlNewProperty.xml)) = oeRecordNotFound Then
            'BMIDS00563 End
                      
                Set xmlTempElement = m_objXmlAssist.GetNode(xmlAccount, ".//PROPERTYADDRESS")
                Set xmlTempElement = xmlTempElement.cloneNode(True)
                xmlTableElem.appendChild xmlTempElement
                m_objXmlAssist.ChangeNodeName xmlTempElement, "PROPERTYADDRESS", "ADDRESS"
                Set xmlTempElement = objNewPropertyBO.CreateNewPropertyAddress(xmlTempRequestDoc.documentElement)
                m_objErrAssist.CheckResponse xmlTempElement, xmlResponseElem, True
            
            End If  'BMIDS00563 End
          End If  ' MAH 02/04/2007 EP2_1597
        End If
        'MV - Start
        Set xmlSpecialFeaturesNodes = xmlAccount.selectNodes(".//SPECIALFEATURE")
        
        If xmlSpecialFeaturesNodes.length > 0 Then
            
            For Each xmlSpecialFeature In xmlSpecialFeaturesNodes
               
                If Len(xmlGetNodeText(xmlSpecialFeature, "SPECIALFEATUREINDICATOR")) > 0 Then
                    
                    Set xmlRequestElem = m_objXmlAssist.GetRequestNodeEx(vxmlRequest)
                    Set xmlTempRequestDoc.documentElement = xmlRequestElem
    
                    Set xmlTableElem = xmlTempRequestDoc.createElement("MORTGAGEACCOUNTSPECIALFEATURE")
                    xmlRequestElem.appendChild xmlTableElem
                    
                    Set xmlAccountGuid = xmlTempRequestDoc.createElement("ACCOUNTGUID")
                    xmlAccountGuid.Text = strAccountGUID
                    xmlTableElem.appendChild xmlAccountGuid
                
                    'BM0358 GHun 11/06/2003 Rename nodes to name expected by ClassDef
                    Set xmlTempNode = xmlSpecialFeature.selectSingleNode("SPECIALFEATUREINDICATOR")
                    xmlChangeNodeName xmlTempNode, "SPECIALFEATUREINDICATOR", "MORTGAGEACCOUNTSPECIALFEATUREIND"
                    xmlTableElem.appendChild xmlTempNode
                    
                    Set xmlTempNode = xmlSpecialFeature.selectSingleNode("SPECIALFEATUREDESCRIPTION")
                    xmlChangeNodeName xmlTempNode, "SPECIALFEATUREDESCRIPTION", "MORTGAGEACCOUNTSPECIALFEATUREDESC"
                    xmlTableElem.appendChild xmlTempNode
                    'BM0358 End
                                
                    strResponse = objCustFinancialBO.CreateSpecialFeature(xmlRequestElem.xml)
                    Set xmlTempResponseDoc = m_objXmlAssist.load(strResponse, TypeName(Me), cstrFunctionName)
                    m_objErrAssist.CheckResponse xmlTempResponseDoc.documentElement, xmlResponseElem, True
            
                End If
            
            Next
            
        End If
        'MV End
        
        'SR 16/01/2007 : EP2_858 - Create ApplicationBankBuildingSoc record (along with corresponding ThirdParty record)
        ' PSC 30/01/2007 EP2_1146
        If (blnReturningCustomer) Then
            
            'Create a record in ApplicationBankBuildingSoc only if AccountNumber and Name exist in ODI's response
            '
            ' PSC 30/01/2007 EP2_1114
            If (xmlGetNodeText(xmlAccount, "BANKACCOUNTNAME") <> "" And xmlGetNodeText(xmlAccount, "BANKACCOUNTNUMBER") <> "") Then
                
                ' PSC 30/01/2007 EP2_1114 - Start
                ' Check if a repayment bank account exists already
                Set xmlRequestElem = m_objXmlAssist.GetRequestNodeEx(vxmlRequest)
                Set xmlTableElem = xmlTempRequestDoc.createElement("APPLICATIONBANKBUILDINGSOC")
                xmlRequestElem.appendChild xmlTableElem
                xmlTableElem.appendChild xmlAppNo.cloneNode(True)
                xmlTableElem.appendChild xmlAppFFNo.cloneNode(True)
                
                Set xmlTempElement = xmlTempRequestDoc.createElement("REPAYMENTBANKACCOUNTINDICATOR")
                xmlTempElement.Text = "1"
                xmlTableElem.appendChild xmlTempElement
                
                strResponse = objApplicationTPBO.FindBankBuildingSocietyList(xmlRequestElem.xml)
                Set xmlTempResponseDoc = m_objXmlAssist.load(strResponse, TypeName(Me), cstrFunctionName)
                 
                lngErrorNumber = m_objErrAssist.GetErrorNumberFromResponse(xmlTempResponseDoc.documentElement, True)
                
                If lngErrorNumber <> 0 And lngErrorNumber <> oeRecordNotFound Then
                    m_objErrAssist.ReRaiseResponseError (xmlTempResponseDoc.documentElement)
                End If
                
                If lngErrorNumber = oeRecordNotFound Then
                
                    Set xmlRequestElem = m_objXmlAssist.GetRequestNodeEx(vxmlRequest)
                    Set xmlTableElem = xmlTempRequestDoc.createElement("APPLICATIONBANKBUILDINGSOC")
                    xmlRequestElem.appendChild xmlTableElem
                    xmlTableElem.appendChild xmlAppNo.cloneNode(True)
                    xmlTableElem.appendChild xmlAppFFNo.cloneNode(True)
                    
                    ' PSC 30/01/2007 EP2_1114 - Start
                    Set xmlTempElement = xmlTempRequestDoc.createElement("ACCOUNTNAME")
                    xmlTempElement.Text = xmlGetNodeText(xmlAccount, "BANKACCOUNTNAME")
                    xmlTableElem.appendChild xmlTempElement
                    
                    ' PSC 30/01/2007 EP2_1114 - Start
                    Set xmlTempElement = xmlTempRequestDoc.createElement("ACCOUNTNUMBER")
                    xmlTempElement.Text = xmlGetNodeText(xmlAccount, "BANKACCOUNTNUMBER")
                    ' PSC 30/01/2007 EP2_1114 - End
                    xmlTableElem.appendChild xmlTempElement
                    
                    ' PSC 30/01/2007 EP2_1114 - Start
                    Set xmlTempElement = xmlTempRequestDoc.createElement("PREFEREDPAYMENTDAY")
                    xmlTempElement.Text = xmlGetNodeText(xmlAccount, "PAYMENTDUEDATE")
                    xmlTableElem.appendChild xmlTempElement
                    
                    Set xmlTempElement = xmlTempRequestDoc.createElement("DDEXPLAINEDIND")
                    xmlTempElement.Text = "1"
                    xmlTableElem.appendChild xmlTempElement
                    
                    Set xmlTempElement = xmlTempRequestDoc.createElement("PROPOSEDPAYMENTMETHOD")
                    xmlTempElement.Text = GetFirstComboValueId("ProposedRepayMethod", "D")
                    xmlTableElem.appendChild xmlTempElement
                    ' PSC 30/01/2007 EP2_1114 - End
                    
                    Set xmlTempElement = xmlTempRequestDoc.createElement("REPAYMENTBANKACCOUNTINDICATOR")
                    xmlTempElement.Text = "1"
                    xmlTableElem.appendChild xmlTempElement
                    
                    Set xmlTempElement = xmlTempRequestDoc.createElement("DIRECTORYADDRESSINDICATOR")
                    xmlTempElement.Text = "0"
                    xmlTableElem.appendChild xmlTempElement
                    
                    Set xmlElement = xmlTempRequestDoc.createElement("THIRDPARTY")
                    xmlTableElem.appendChild xmlElement
                    
                    Set xmlTempElement = xmlTempRequestDoc.createElement("THIRDPARTYTYPE")
                    xmlTempElement.Text = "3"
                    xmlElement.appendChild xmlTempElement
                    
                    Set xmlTempElement = xmlTempRequestDoc.createElement("COMPANYNAME")
                    xmlTempElement.Text = xmlGetNodeText(xmlAccount, "BANKNAME")
                    xmlElement.appendChild xmlTempElement
                    
                    ' PSC 30/01/2007 EP2_1114 - Start
                    strSortCode = xmlGetNodeText(xmlAccount, "BANKSORTCODE")
                    
                    If Len(strSortCode) = 6 Then
                        nDashPos = InStr(1, strSortCode, "-", vbTextCompare)
                        If nDashPos = 0 Then
                            strSortCode = Left(strSortCode, 2) + "-" + Mid(strSortCode, 3, 2) + "-" + Right(strSortCode, 2)
                        End If
                    End If
    
                    Set xmlTempElement = xmlTempRequestDoc.createElement("THIRDPARTYBANKSORTCODE")
                    xmlTempElement.Text = strSortCode
                    ' PSC 30/01/2007 EP2_1114 - End
                    
                    xmlElement.appendChild xmlTempElement
                    
                    strResponse = objApplicationTPBO.CreateBankBuildingSociety(xmlRequestElem.xml)
                    Set xmlTempResponseDoc = m_objXmlAssist.load(strResponse, TypeName(Me), cstrFunctionName)
                    m_objErrAssist.CheckResponse xmlTempResponseDoc.documentElement, xmlResponseElem, True
                End If
                ' PSC 30/01/2007 EP2_1114 - End

            End If
        End If
        'SR 16/01/2007 : EP2_858 - End
    Next
                  
    'SR 13/08/2004 : BMIDS815
    Set xmlTableElem = xmlTempRequestDoc.createElement("APPLICATION")
    xmlTableElem.appendChild xmlAppNo.cloneNode(True)
    xmlTableElem.appendChild xmlAppFFNo.cloneNode(True)
    
    Set xmlTempElement = xmlTempRequestDoc.createElement("LASTACCOUNTDOWNLOADDATE")
    xmlTempElement.Text = Now()
    xmlTableElem.appendChild xmlTempElement
    
    objIApplicationDO.Update xmlTableElem
    'SR 13/08/2004 : BMIDS815 - End
    
    Set IApplicationManagerTxBO_ImportAccountsIntoApplication = xmlResponseElem
    
    m_objContext.SetComplete
    
ImportAccountsIntoApplicationExit:

    Set xmlElement = Nothing
    Set xmlOut = Nothing
    Set xmlResponseElem = Nothing
    Set objIApplicationBO = Nothing
    Set xmlAppNo = Nothing
    Set xmlAppFFNo = Nothing
    Set objCustFinancialBO = Nothing
    Set xmlTempRequestDoc = Nothing
    Set xmlRequestElem = Nothing
    Set xmlTableElem = Nothing
    Set xmlTempResponseDoc = Nothing
    Set xmlAccount = Nothing
    Set xmlMortgageAccountDoc = Nothing
    'Set xmlListElem = Nothing
    Set xmlTempElement = Nothing
    Set xmlAccountList = Nothing
    'Set objConvAssist = Nothing
    Set objNewPropertyBO = Nothing
    Set xmlAccountGuid = Nothing
    Set xmlOtherSystemAccountNumber = Nothing
    Set xmlElemBasedAdmin = Nothing

    Set xmlPropertyNode = Nothing

    'BMIDS00563
    Set xmlNewProperty = Nothing
    'BMIDS00563 End

    Set objIApplicationDO = Nothing  'SR 13/08/2004 : BMIDS815
    Set objApplicationTPBO = Nothing
    Set xmlExistingFS = Nothing 'BM0377

    'BM0387
    Set xmlArrearsHistoryList = Nothing
    Set xmlAccountArrears = Nothing
    'BM0387 End
    
    
    Exit Function

ImportAccountsIntoApplicationVbErr:

    If m_objErrAssist.GetOmigaErrorNumber(Err.Number) = oeRecordNotFound Then
        Resume Next
    End If
    
    If m_objErrAssist.IsWarning = True Then
        m_objErrAssist.AddWarning xmlResponseElem
        Resume Next
    End If

    m_objErrAssist.AddToErrSource cstrFunctionName
    
    If m_objErrAssist.IsSystemError = True Then
        m_objErrAssist.LogError TypeName(Me), cstrFunctionName, Err.Number, Err.Description
    End If

    Set IApplicationManagerTxBO_ImportAccountsIntoApplication = m_objErrAssist.CreateErrorResponseEx()

    m_objContext.SetAbort
    
    ' go to clean-up section
    Resume ImportAccountsIntoApplicationExit

End Function

Private Function CreateLocks(ByVal vxmlRequest As IXMLDOMElement, _
                           ByVal vstrLockType As String, _
                           ByVal vstrKeyNumber As String) As IXMLDOMElement
' header ---------------------------------------------------------------------------------
' description:  Creates a lock based on Lock Type passed in
' pass:
'   vstrXmlRequest
'       Xml Request containing a valid request tag:
'   vstrLockType
'       The type of lock to create
'   vstrKeyNumber
'       The key to apply to the lock record
' return: n/a
'------------------------------------------------------------------------------------------
On Error GoTo CreateLocksVbErr

    Const strFunctionName As String = "CreateLocks"

    Dim xmlOut As New FreeThreadedDOMDocument40
    xmlOut.validateOnParse = False
    xmlOut.setProperty "NewParser", True
    Dim xmlResponseElem As IXMLDOMElement, xmlRequestElem As IXMLDOMElement
    Dim xmlTableElem As IXMLDOMElement, xmlTempResponseNode As IXMLDOMElement
    Dim xmlElement As IXMLDOMElement

    Dim objICustomerBO As ICustomerBO, objIApplManagerBO As IApplicationManagerBO

    Dim strTableName As String, strTypeNumber As String

    Set xmlResponseElem = xmlOut.createElement("RESPONSE")
    xmlOut.appendChild xmlResponseElem
    xmlResponseElem.setAttribute "TYPE", "SUCCESS"

    ' Setup a create request
    strTableName = UCase(vstrLockType) & "LOCK"
    strTypeNumber = UCase(vstrLockType) & "NUMBER"

    If vxmlRequest.nodeName = "REQUEST" Then
        Set xmlRequestElem = vxmlRequest.cloneNode(False)
    Else
        ' Raise an error
    End If

    Set xmlTableElem = xmlOut.createElement(strTableName)
    xmlRequestElem.appendChild xmlTableElem

    Set xmlElement = xmlOut.createElement(strTypeNumber)
    xmlElement.Text = vstrKeyNumber
    xmlTableElem.appendChild xmlElement

    Set xmlElement = xmlOut.createElement("UNITID")
    xmlElement.Text = _
        m_objXmlAssist.GetAttributeValue(vxmlRequest, "REQUEST", "UNITID")
    xmlTableElem.appendChild xmlElement

    Set xmlElement = xmlOut.createElement("USERID")
    xmlElement.Text = _
        m_objXmlAssist.GetAttributeValue(vxmlRequest, "REQUEST", "USERID")
    xmlTableElem.appendChild xmlElement
    Set xmlElement = xmlOut.createElement("MACHINEID")
    xmlElement.Text = _
        m_objXmlAssist.GetAttributeValue(vxmlRequest, "REQUEST", "MACHINEID")
    xmlTableElem.appendChild xmlElement
    Set xmlElement = xmlOut.createElement("TYPEOFLOCK")
    xmlElement.Text = _
        m_objXmlAssist.GetAttributeValue(vxmlRequest, "REQUEST", "TYPEOFLOCK")
    xmlTableElem.appendChild xmlElement
    
    Set xmlElement = xmlOut.createElement("CHANNELID")
    xmlElement.Text = _
        m_objXmlAssist.GetAttributeValue(vxmlRequest, "REQUEST", "CHANNELID")
    xmlTableElem.appendChild xmlElement
    

    ' Call the appropriate Business Object
    If UCase(vstrLockType) = "CUSTOMER" Then
        If m_objContext Is Nothing Then
            Set objICustomerBO = New CustomerBO
        Else
            Set objICustomerBO = _
                m_objContext.CreateInstance(gstrCUSTOMER_COMPONENT & ".CustomerBO")
        End If
        Set xmlTempResponseNode = xmlOut.createElement("RESPONSE")
        objICustomerBO.CreateCustomerLock xmlRequestElem, xmlTempResponseNode, "1"
        m_objErrAssist.CheckResponse xmlTempResponseNode, xmlResponseElem, True

    ElseIf UCase(vstrLockType) = "APPLICATION" Then

        ' RF 02/12/99 AQR SC017: Enhancement -
        ' Validate working hours on creating an application lock.
        Set xmlElement = xmlOut.createElement("CHANNELID")
        xmlElement.Text = m_objXmlAssist.GetTagValue(vxmlRequest, "CHANNELID")
        xmlTableElem.appendChild xmlElement

         If m_objContext Is Nothing Then
            Set objIApplManagerBO = New ApplicationManagerBO
        Else
            Set objIApplManagerBO = _
                m_objContext.CreateInstance(App.Title & ".ApplicationManagerBO")
        End If

        Set xmlTempResponseNode = objIApplManagerBO.Validate(xmlRequestElem, bomCreateApplicationLock)
        m_objErrAssist.CheckResponse xmlTempResponseNode, xmlResponseElem, True
        
        'SYS0611 MC 06/06/2000. Add lock type and date when locking existing application
        AddLockTypeAndDate xmlRequestElem
        
        Set xmlTempResponseNode = IApplicationManagerTxBO_CreateLock(xmlRequestElem)
        m_objErrAssist.CheckResponse xmlTempResponseNode, xmlResponseElem, True
    End If

    Set CreateLocks = xmlResponseElem

    If Not m_objContext Is Nothing Then
        m_objContext.SetComplete
    End If
    
CreateLocksExit:
   
    Set xmlOut = Nothing
    Set xmlResponseElem = Nothing
    Set xmlTempResponseNode = Nothing
    Set xmlRequestElem = Nothing
    Set xmlTableElem = Nothing
    Set xmlElement = Nothing
    
    Set objICustomerBO = Nothing
    Set objIApplManagerBO = Nothing

    Exit Function

CreateLocksVbErr:
    
    If m_objErrAssist.IsWarning = True Then
        m_objErrAssist.AddWarning xmlResponseElem
        Resume Next
    End If

    m_objErrAssist.AddToErrSource strFunctionName
    
    If m_objErrAssist.IsSystemError = True Then
        m_objErrAssist.LogError TypeName(Me), strFunctionName, Err.Number, Err.Description
    End If

    Set CreateLocks = m_objErrAssist.CreateErrorResponseEx()

    If Not m_objContext Is Nothing Then
        m_objContext.SetAbort
    End If
    
    ' go to clean-up section
    Resume CreateLocksExit

End Function

Private Sub CreateCustLockApplLock(ByVal objCustomerRoleElem As IXMLDOMElement)
' header --------------------------------------------------------------
' description:  Creates a CustomerLockApplicationLock for the customer
' pass:         objCustomerRoleElem  Customer role element
'
' return:       n/a
'
' Raise Errors:
'----------------------------------------------------------------------
On Error GoTo CreateLockVbErr

    Const strFunctionName As String = "CreateCustLockApplLock"

    Dim objRequestDoc As New FreeThreadedDOMDocument40
    objRequestDoc.validateOnParse = False
    objRequestDoc.setProperty "NewParser", True
    Dim objTableElem As IXMLDOMElement
    Dim objRequestNode As IXMLDOMNode
    Dim objElement As IXMLDOMElement
    Dim objIApplManagerDO As IApplicationManagerDO
    Dim objGlobalParamDO As GlobalParameterDO

    If objCustomerRoleElem Is Nothing Then
        m_objErrAssist.ThrowError TypeName(Me), strFunctionName, _
                                oeInvalidParameter, _
                                "Customer Role element is set to nothing"
    End If

    Dim strApplicationNo As String
    Dim strCustomerNo As String

    strApplicationNo = m_objXmlAssist.GetTagValue(objCustomerRoleElem, "APPLICATIONNUMBER")
    strCustomerNo = m_objXmlAssist.GetTagValue(objCustomerRoleElem, "CUSTOMERNUMBER")

    If m_objContext Is Nothing Then
        Set objIApplManagerDO = New ApplicationManagerDO
        Set objGlobalParamDO = New GlobalParameterDO
    Else
        Set objIApplManagerDO = m_objContext.CreateInstance(cstrCOMPONENT_NAME & _
                                                             ".ApplicationManagerDO")
        Set objGlobalParamDO = m_objContext.CreateInstance(cstrBASE & _
                                                             ".GlobalParameterDO")
    End If

    'decide if we want a lock or not
    Dim strLockingReqd As String
    strLockingReqd = objGlobalParamDO.GetCurrentParameterByType("Locking", "BOOLEAN")
    If strLockingReqd = "1" Then
        ' Set up create request for CustomerLockApplicationLock
    
        Set objRequestNode = m_objXmlAssist.GetRequestNode(objCustomerRoleElem.ownerDocument)
        objRequestDoc.appendChild objRequestNode
        Set objTableElem = objRequestDoc.createElement("CUSTOMERLOCKAPPLICATIONLOCK")
        objRequestNode.appendChild objTableElem
        Set objElement = objRequestDoc.createElement("APPLICATIONNUMBER")
        objElement.Text = strApplicationNo
        objTableElem.appendChild objElement
        Set objElement = objRequestDoc.createElement("CUSTOMERNUMBER")
        objElement.Text = strCustomerNo
        objTableElem.appendChild objElement
    
        objIApplManagerDO.CreateCustomerLockApplicationLock objRequestDoc.documentElement
    End If
      
    Set objRequestDoc = Nothing
    Set objTableElem = Nothing
    Set objRequestNode = Nothing
    Set objElement = Nothing
    Set objIApplManagerDO = Nothing
    Set objGlobalParamDO = Nothing

    Exit Sub

CreateLockVbErr:

    Set objRequestDoc = Nothing
    Set objTableElem = Nothing
    Set objRequestNode = Nothing
    Set objElement = Nothing
    Set objIApplManagerDO = Nothing
    Set objGlobalParamDO = Nothing

    ' Re-raise error for business object to interpret as appropriate
    Err.Raise Err.Number, Err.Source, Err.Description

End Sub

Private Sub SetInputForCreateFullApplication( _
                ByVal vxmlRequest As IXMLDOMElement, _
                ByVal strTableName As String, _
                ByVal strApplNumber As String, ByVal strAFFNumber As String, _
                Optional ByVal bErrorIfTagNotFound As Boolean = True)

On Error GoTo ErrHandle
    
    Const strFunctionName As String = "SetInputForCreateFullApplication"
    
    Dim xmlElement As IXMLDOMElement
    Dim xmlApplElem As IXMLDOMElement
    Dim xmlAFFElem As IXMLDOMElement
    Dim xmlOut As New FreeThreadedDOMDocument40
    xmlOut.validateOnParse = False
    xmlOut.setProperty "NewParser", True
    Dim xmlNodeList As IXMLDOMNodeList
    
    Set xmlNodeList = vxmlRequest.selectNodes(".//" & strTableName)
    If xmlNodeList.length = 0 Then
        If bErrorIfTagNotFound Then
            m_objErrAssist.ThrowError TypeName(Me), strFunctionName, _
                                oeMissingPrimaryTag, strTableName & " tag not found"
        Else
            'Exit subroutine without raising an error
            Exit Sub
        End If
    End If
    
    For Each xmlElement In xmlNodeList
        
        Set xmlApplElem = xmlElement.selectSingleNode(".//APPLICATIONNUMBER")
        If xmlApplElem Is Nothing Then
            Set xmlApplElem = xmlOut.createElement("APPLICATIONNUMBER")
        End If
        xmlApplElem.Text = strApplNumber
        
        Set xmlAFFElem = xmlElement.selectSingleNode(".//APPLICATIONFACTFINDNUMBER")
        If xmlAFFElem Is Nothing Then
            Set xmlAFFElem = xmlOut.createElement("APPLICATIONFACTFINDNUMBER")
        End If
        xmlAFFElem.Text = strAFFNumber
    
    Next
    
    Set xmlElement = Nothing
    Set xmlApplElem = Nothing
    Set xmlAFFElem = Nothing
    Set xmlOut = Nothing
    Exit Sub

ErrHandle:
    Set xmlElement = Nothing
    Set xmlApplElem = Nothing
    Set xmlAFFElem = Nothing
    Set xmlOut = Nothing
    
    Err.Raise Err.Number, Err.Source, Err.Description
End Sub

Private Function GetUserHodingCustomerOrApplication( _
                            ByVal vxmlRequest As IXMLDOMElement, _
                            ByVal strLockType As String) As String
'--------------------------------------------------------------------------------
'Description : Finds the Id of the user holding the lock either Customer or
'              Application.
'Pass        : vxmlRequest - node containing the Application/Customer data
'              strLockType - Type of Lock (Application or Customer)
'Return      : UserId
'--------------------------------------------------------------------------------

On Error GoTo ErrHandle
    
    Const strFunctionName As String = "GetUserHodingCustomerOrApplication"
    
    Dim strNumber As String
    Dim objIADOAssist As IADOAssist
    
    Dim strCondition As String
    Dim strUserId As String
    
    If m_objContext Is Nothing Then
        Set objIADOAssist = New ADOAssist
    Else
        Set objIADOAssist = m_objContext.CreateInstance(gstrBASE_COMPONENT & ".ADOAssist")
    End If
    
    If strLockType = "APPLICATION" Then
        strNumber = vxmlRequest.selectSingleNode(".//APPLICATIONNUMBER").Text
        strCondition = " APPLICATIONNUMBER = '" & strNumber & "'"
        ' Return empty string, if the rcord was not found in the table
        If IsNull(objIADOAssist.GetValueFromTable("APPLICATIONLOCK", strCondition, "USERID", strUserId)) Then
            strUserId = ""
        End If
    Else
        strNumber = vxmlRequest.selectSingleNode(".//CUSTOMERNUMBER").Text
        strCondition = " CUSTOMERNUMBER = '" & strNumber & "'"
        ' Return empty string, if the rcord was not found in the table
        If IsNull(objIADOAssist.GetValueFromTable("CUSTOMERLOCK", strCondition, "USERID", strUserId)) Then
            strUserId = ""
        End If
    End If
    
    GetUserHodingCustomerOrApplication = strUserId
    
    Set objIADOAssist = Nothing
    
    If Not m_objContext Is Nothing Then
        m_objContext.SetComplete
    End If
    
    Exit Function
ErrHandle:
    Set objIADOAssist = Nothing
    m_objErrAssist.AddToErrSource strFunctionName
    
    Err.Raise Err.Number, Err.Source, Err.Description
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

Private Function ImportCustomers(ByVal vxmlRequest As IXMLDOMElement, _
                                 ByVal vxmlMortgageAccount As IXMLDOMElement, _
                                 ByVal vstrApplicationNumber As String, _
                                 ByVal vstrApplicationFFNumber As String, _
                                 ByVal vblnAddToApplication As Boolean, _
                                 ByVal vblnAddAllCustomers As Boolean, _
                                 ByVal vxmlResponse As IXMLDOMElement, _
                                 ByRef rstrAdditionalDetails As String) As Boolean
' header ----------------------------------------------------------------------------------
' description:  Creates the additional details for the applicants attached to an account.
'               Optionally checks to see if the customers already exist in Omiga and if not
'               creates them and linkds them to the application
' pass:         vxmlRequest             xml Request
'               vxmlMortgageAccount     Mortgage Account containing the customers
'               vstrApplicationNumber   Application to attach customers to
'               vstrApplicationFFNumber Application Fact Find number of application to attach
'                                       customers to
'               vblnAddToApplication    Denotes whether to add the customers to the application
'               vblnAddAllCustomers     Denotes whether to add all customers or just the one
'                                       passed in
'               vxmlResponse            xml Response to add warnings to
'               rstrAdditionalDetails   The additional details for the customers on the account
                                        

' return:       ImportCustomers         Whether there are additional details or not
'------------------------------------------------------------------------------------------
On Error GoTo ImportCustomersVbErr

    Const cstrFunctionName As String = "ImportCustomers"
    

    Dim objCustomerBO As CustomerBO
    Dim objICustomerRoleBO As ICustomerRoleBO
    Dim objComboDO As ComboDO
    Dim xmlCustomerList As IXMLDOMNodeList
    Dim xmlCustomer As IXMLDOMElement
    Dim xmlCustRoleDoc As FreeThreadedDOMDocument40
    Dim xmlTempResponseDoc As FreeThreadedDOMDocument40
    Dim xmlTableElem As IXMLDOMElement
    Dim xmlElement As IXMLDOMElement
    Dim xmlRequestElem As IXMLDOMElement
    Dim xmlCustomerNo As IXMLDOMElement
    Dim xmlCustomerVersionNo As IXMLDOMElement
    Dim xmlCustomerRoleType As IXMLDOMElement
    Dim xmlCustomerOrder As IXMLDOMElement
    Dim xmlOtherSysCustNo As IXMLDOMElement
    Dim xmlNewCustNo As IXMLDOMElement
    Dim xmlNewCustVersionNo As IXMLDOMElement
    Dim xmlFindCustomerDoc As FreeThreadedDOMDocument40
    Dim xmlFindCustomerNo As IXMLDOMNode                ' PSC 03/10/03 BMIDS612
    Dim xmlCustomerDoc As FreeThreadedDOMDocument40
    Dim xmlCustRequestElem As IXMLDOMElement
    Dim xmlConvertedCustElem As IXMLDOMElement
    Dim xmlActionElem As IXMLDOMElement
    Dim xmlCustLockDoc As FreeThreadedDOMDocument40
    Dim xmlLockCustNoElem As IXMLDOMElement
    
    'BMIDS00004 MDC 21/05/2002 - Versions of Employment
    Dim xmlTempRequest As IXMLDOMNode
    Dim xmlSearchNode As IXMLDOMNode
    Dim xmlCustomerNode As IXMLDOMNode
    Dim strNewCustVersionNumber As String
    'BMIDS00004 MDC 21/05/2002 - End
    
    Dim strCustomerName As String
    Dim blnOtherCustomers As Boolean
    Dim strPattern As String
    Dim strOtherSysCustNo As String
    Dim strResponse As String
    Dim strSurname As String
    Dim strFirstName As String
    Dim strSecondName As String
    Dim strOtherNames As String
    Dim intAFFNumber As Integer     'BMIDS00566
    
    Set objCustomerBO = m_objContext.CreateInstance(gstrCUSTOMER_COMPONENT & ".CustomerBO")
    Set objComboDO = m_objContext.CreateInstance(gstrBASE_COMPONENT & ".ComboDO")
    Set objICustomerRoleBO = m_objContext.CreateInstance(App.Title & ".CustomerRoleBO")
        
    strOtherSysCustNo = m_objXmlAssist.GetNodeValue(vxmlRequest, ".//OTHERSYSTEMCUSTOMERNUMBER")
    
    blnOtherCustomers = False
    strPattern = ".//CUSTOMER[CUSTOMERNUMBER != '" & strOtherSysCustNo & "']"
    Set xmlCustomerList = vxmlMortgageAccount.selectNodes(strPattern)
    
    If xmlCustomerList.length > 0 Then
        blnOtherCustomers = True
    End If
            
    rstrAdditionalDetails = ""
    
    ' Create customer role request and customer request if we are adding the
    ' customer(s) to the application
    If vblnAddToApplication = True Then
        Set xmlCustRoleDoc = New FreeThreadedDOMDocument40
        xmlCustRoleDoc.validateOnParse = False
        xmlCustRoleDoc.setProperty "NewParser", True
        xmlCustRoleDoc.async = False
        Set xmlRequestElem = m_objXmlAssist.GetRequestNodeEx(vxmlRequest)
        xmlCustRoleDoc.appendChild xmlRequestElem
        Set xmlTableElem = xmlCustRoleDoc.createElement("CUSTOMERROLE")
        xmlRequestElem.appendChild xmlTableElem
        Set xmlCustomerNo = xmlCustRoleDoc.createElement("CUSTOMERNUMBER")
        xmlTableElem.appendChild xmlCustomerNo
        Set xmlCustomerVersionNo = xmlCustRoleDoc.createElement("CUSTOMERVERSIONNUMBER")
        xmlTableElem.appendChild xmlCustomerVersionNo
        Set xmlElement = xmlCustRoleDoc.createElement("APPLICATIONNUMBER")
        xmlElement.Text = vstrApplicationNumber
        xmlTableElem.appendChild xmlElement
        Set xmlElement = xmlCustRoleDoc.createElement("APPLICATIONFACTFINDNUMBER")
        xmlElement.Text = vstrApplicationFFNumber
        xmlTableElem.appendChild xmlElement
        Set xmlCustomerRoleType = xmlCustRoleDoc.createElement("CUSTOMERROLETYPE")
        xmlTableElem.appendChild xmlCustomerRoleType
        Set xmlCustomerOrder = xmlCustRoleDoc.createElement("CUSTOMERORDER")
        xmlTableElem.appendChild xmlCustomerOrder
            
        Set xmlFindCustomerDoc = New FreeThreadedDOMDocument40
        xmlFindCustomerDoc.validateOnParse = False
        xmlFindCustomerDoc.setProperty "NewParser", True
        xmlFindCustomerDoc.async = False
        Set xmlRequestElem = m_objXmlAssist.GetRequestNodeEx(vxmlRequest)
        xmlFindCustomerDoc.appendChild xmlRequestElem
        Set xmlTableElem = xmlFindCustomerDoc.createElement("CUSTOMER")
        xmlRequestElem.appendChild xmlTableElem
        Set xmlOtherSysCustNo = xmlFindCustomerDoc.createElement("OTHERSYSTEMCUSTOMERNUMBER")
        xmlTableElem.appendChild xmlOtherSysCustNo
        
        ' PSC 03/10/03 BMIDS612 - Start
        Set xmlFindCustomerNo = xmlFindCustomerDoc.createElement("CUSTOMERNUMBER")
        xmlTableElem.appendChild xmlFindCustomerNo
        ' PSC 03/10/03 BMIDS612 - End
        
        Set xmlCustomerDoc = New FreeThreadedDOMDocument40
        xmlCustomerDoc.validateOnParse = False
        xmlCustomerDoc.setProperty "NewParser", True
        xmlCustomerDoc.async = False
        Set xmlCustRequestElem = m_objXmlAssist.GetRequestNodeEx(vxmlRequest)
        xmlCustomerDoc.appendChild xmlCustRequestElem
        Set xmlActionElem = xmlCustomerDoc.createElement("CREATE")
        xmlCustRequestElem.appendChild xmlActionElem
        
        Set xmlCustLockDoc = New FreeThreadedDOMDocument40
        xmlCustLockDoc.validateOnParse = False
        xmlCustLockDoc.setProperty "NewParser", True
        xmlCustLockDoc.async = False
        Set xmlRequestElem = m_objXmlAssist.GetRequestNodeEx(vxmlRequest)
        xmlCustLockDoc.appendChild xmlRequestElem
        Set xmlTableElem = xmlCustLockDoc.createElement("CUSTOMER")
        xmlRequestElem.appendChild xmlTableElem
        Set xmlLockCustNoElem = xmlCustLockDoc.createElement("CUSTOMERNUMBER")
        xmlTableElem.appendChild xmlLockCustNoElem
    
        'BMIDS00004 MDC 18/07/2002 - Versions of Employment
        'Create new customer version for the new application
        Set xmlTempRequest = vxmlRequest.cloneNode(False)
        Set xmlSearchNode = xmlCustomerDoc.createElement("SEARCH")
        xmlTempRequest.appendChild xmlSearchNode
        Set xmlCustomerNode = xmlCustomerDoc.createElement("CUSTOMER")
        xmlSearchNode.appendChild xmlCustomerNode
        xmlCustomerNode.appendChild m_objXmlAssist.GetMandatoryNode(vxmlRequest, ".//CUSTOMERNUMBER").cloneNode(True)
        xmlCustomerNode.appendChild m_objXmlAssist.GetMandatoryNode(vxmlRequest, ".//CUSTOMERVERSIONNUMBER").cloneNode(True)
        
        strResponse = objCustomerBO.CreateNewCustomerVersion(xmlTempRequest.xml)
        Set xmlTempResponseDoc = m_objXmlAssist.load(strResponse, TypeName(Me), cstrFunctionName)
        m_objErrAssist.CheckResponse xmlTempResponseDoc.documentElement, vxmlResponse, True
                                                                     
        Set xmlNewCustVersionNo = m_objXmlAssist.GetNode(xmlTempResponseDoc.documentElement, _
                                                        ".//CUSTOMERVERSIONNUMBER")
        If Not xmlNewCustVersionNo Is Nothing Then
            strNewCustVersionNumber = xmlNewCustVersionNo.Text
        End If
        'BMIDS00004 MDC 18/07/2002 - End
    
    End If

    
    ' For each additional customer add their details into the additional details.
    ' If they are to be added to the application check if they exist and create
    ' them if they don't. Then link them to the application
    For Each xmlCustomer In xmlCustomerList
            
        strCustomerName = ""
        
        ' Get this customer's name
        strFirstName = m_objXmlAssist.GetNodeValue(xmlCustomer, _
                                                      ".//FIRSTFORENAME", _
                                                      False, False)
        
        strSecondName = strCustomerName & _
                            m_objXmlAssist.GetNodeValue(xmlCustomer, _
                                                        ".//SECONDFORENAME", _
                                                        False, False)
        
        strOtherNames = strCustomerName & _
                            m_objXmlAssist.GetNodeValue(xmlCustomer, _
                                                        ".//OTHERFORENAMES", _
                                                        False, False)
        
        strSurname = strCustomerName & _
                            m_objXmlAssist.GetNodeValue(xmlCustomer, _
                                                        ".//SURNAME", _
                                                        False, False)
        strCustomerName = strFirstName
        
        If Len(strCustomerName) > 0 And Len(strSecondName) > 0 Then
            strCustomerName = strCustomerName & " "
        End If
        
        strCustomerName = strCustomerName & strSecondName
        
        If Len(strCustomerName) > 0 And Len(strOtherNames) > 0 Then
            strCustomerName = strCustomerName & " "
        End If
        
        strCustomerName = strCustomerName & strOtherNames
        
        If Len(strCustomerName) > 0 And Len(strSurname) > 0 Then
            strCustomerName = strCustomerName & " "
        End If
        
        strCustomerName = strCustomerName & strSurname & vbCr
        
        ' Append it to the details of all the customers on this account
        rstrAdditionalDetails = rstrAdditionalDetails & strCustomerName
    
        ' If adding customers from the application and we need to add all customers
        If vblnAddToApplication = True And vblnAddAllCustomers = True Then
                
             ' See if this customer exists in Omiga
             xmlOtherSysCustNo.Text = m_objXmlAssist.GetNodeValue(xmlCustomer, ".//CUSTOMERNUMBER")
             ' PSC 03/10/03 BMIDS612
             ' PSC 15/10/03 BMIDS655
             xmlFindCustomerNo.Text = m_objXmlAssist.GetNodeValue(xmlCustomer, ".//OMIGACUSTOMERNUMBER", False)
             strResponse = objCustomerBO.GetCustomerFromLegacyNumber(xmlFindCustomerDoc.xml)
             Set xmlTempResponseDoc = m_objXmlAssist.load(strResponse, _
                                                          TypeName(Me), _
                                                          cstrFunctionName)
             m_objErrAssist.CheckResponse xmlTempResponseDoc.documentElement, vxmlResponse, True
            
             Set xmlNewCustNo = _
                     m_objXmlAssist.GetNode(xmlTempResponseDoc.documentElement, _
                                            ".//CUSTOMERNUMBER", False)
             
             Set xmlNewCustVersionNo = _
                     m_objXmlAssist.GetNode(xmlTempResponseDoc.documentElement, _
                                            ".//CUSTOMERVERSIONNUMBER", False)
             
             Set xmlTempResponseDoc = Nothing
                
            ' Create customer if not already there
            If xmlNewCustNo Is Nothing Then
            
                ' Convert the customer to allow creation in Omiga
                Set xmlConvertedCustElem = ConvertCustomerToOmigaFormat(xmlCustomer)
                
                'Replace the details created before hand with these details
                If xmlActionElem.childNodes.length > 0 Then
                    xmlActionElem.replaceChild xmlConvertedCustElem, _
                                               xmlActionElem.firstChild
                Else
                    xmlActionElem.appendChild xmlConvertedCustElem
                End If
                
                strResponse = objCustomerBO.CreateCustomerDetails(xmlCustomerDoc.xml)
                Set xmlTempResponseDoc = m_objXmlAssist.load(strResponse, _
                                                             TypeName(Me), _
                                                             cstrFunctionName)
                m_objErrAssist.CheckResponse xmlTempResponseDoc.documentElement, vxmlResponse, True
                                                                             
                Set xmlNewCustNo = _
                        m_objXmlAssist.GetNode(xmlTempResponseDoc.documentElement, _
                                               ".//CUSTOMERNUMBER")
                Set xmlNewCustVersionNo = _
                        m_objXmlAssist.GetNode(xmlTempResponseDoc.documentElement, _
                                               ".//CUSTOMERVERSIONNUMBER")
                                               
                Set xmlTempResponseDoc = Nothing
                
            Else
                
                ' Lock the existing customer
                xmlLockCustNoElem.Text = xmlNewCustNo.Text
                strResponse = objCustomerBO.CreateCustomerLock(xmlCustLockDoc.xml)
                Set xmlTempResponseDoc = m_objXmlAssist.load(strResponse, _
                                                             TypeName(Me), _
                                                             cstrFunctionName)
                m_objErrAssist.CheckResponse xmlTempResponseDoc.documentElement, vxmlResponse, True
                
                'BMIDS00004 MDC 21/05/2002 - Versions of Employment
                'Create new customer version
                Set xmlTempRequest = vxmlRequest.cloneNode(False)
                Set xmlSearchNode = xmlCustomerDoc.createElement("SEARCH")
                xmlTempRequest.appendChild xmlSearchNode
                Set xmlCustomerNode = xmlCustomerDoc.createElement("CUSTOMER")
                xmlSearchNode.appendChild xmlCustomerNode
                xmlCustomerNode.appendChild xmlNewCustNo
                xmlCustomerNode.appendChild xmlNewCustVersionNo
                
                strResponse = objCustomerBO.CreateNewCustomerVersion(xmlTempRequest.xml)
                Set xmlTempResponseDoc = m_objXmlAssist.load(strResponse, TypeName(Me), cstrFunctionName)
                m_objErrAssist.CheckResponse xmlTempResponseDoc.documentElement, vxmlResponse, True
                                                                             
                Set xmlNewCustVersionNo = m_objXmlAssist.GetNode(xmlTempResponseDoc.documentElement, _
                                                                ".//CUSTOMERVERSIONNUMBER")
                'BMIDS00004 MDC 21/05/2002 - End
                
                Set xmlTempResponseDoc = Nothing
            
            End If
            
            ' Link customer to application
            xmlCustomerNo.Text = xmlNewCustNo.Text
            xmlCustomerVersionNo.Text = xmlNewCustVersionNo.Text
            xmlCustomerRoleType.Text = m_objXmlAssist.GetNodeValue(xmlCustomer, ".//CUSTOMERROLETYPE")
            xmlCustomerOrder.Text = m_objXmlAssist.GetNodeValue(xmlCustomer, ".//CUSTOMERORDER")

            Set xmlElement = objICustomerRoleBO.Create(xmlCustRoleDoc.documentElement)
            m_objErrAssist.CheckResponse xmlElement, vxmlResponse, True
            'BMIDS00566 Core AQR SYS4898 - Add to appliaction salutation if this is one of the first two applicants
            If (CSafeInt(xmlCustomerOrder.Text) <= 2) Then
               intAFFNumber = CSafeInt(vstrApplicationFFNumber)
               UpdateAppCorresSalutation vstrApplicationNumber, intAFFNumber
            End If
            'BMIDS00566 End
        End If
    Next
      
    'SG 04/04/02 SYS4360
    'The code block below may well have been executed in the loop above.
    'Amended If statement so that it's only executed if the corresponding code in the loop above has not.
          
    ' Link the customer passed in to the application
    
    'SG 04/04/02 SYS4360
    'MV - 20/06/2002 - BMIDS0077 - Core Upgrade
    If vblnAddToApplication = True Then
    'If vblnAddToApplication = True And vblnAddAllCustomers = False Then
    
        strPattern = ".//CUSTOMER[CUSTOMERNUMBER = '" & strOtherSysCustNo & "']"
        Set xmlCustomer = m_objXmlAssist.GetNode(vxmlMortgageAccount, strPattern)
        
        xmlCustomerNo.Text = m_objXmlAssist.GetNodeValue(vxmlRequest, ".//CUSTOMERNUMBER")
        
        'BMIDS00004 MDC 18/07/2002 - Customer Versioning
        'Use new version of the customer
        If Len(Trim$(strNewCustVersionNumber)) > 0 Then
            xmlCustomerVersionNo.Text = strNewCustVersionNumber
        Else
            xmlCustomerVersionNo.Text = m_objXmlAssist.GetNodeValue(vxmlRequest, ".//CUSTOMERVERSIONNUMBER")
        End If
        'BMIDS00004 MDC 18/07/2002 - End
        
        'SG 04/04/02 SYS4360 START
        'Amending the If statement above renders the following If statement obsolete
        
        ' If not adding the other customers on then set this one to be the first applicant on
        ' the new application else use the order and type on the account
        'MV - 20/06/2002 - BMIDS0077 - Core Upgrade
        If vblnAddAllCustomers = False Then
            xmlCustomerRoleType.Text = objComboDO.GetFirstComboValueId("CustomerRoleType", "A")
            xmlCustomerOrder.Text = "1"
        Else
            xmlCustomerRoleType.Text = m_objXmlAssist.GetNodeValue(xmlCustomer, ".//CUSTOMERROLETYPE")
            xmlCustomerOrder.Text = m_objXmlAssist.GetNodeValue(xmlCustomer, ".//CUSTOMERORDER")
        End If
        'SG 04/04/02 SYS4360 END
        
        Set xmlElement = objICustomerRoleBO.Create(xmlCustRoleDoc.documentElement)
        m_objErrAssist.CheckResponse xmlElement, vxmlResponse, True
        'BMIDS00566 Core AQR SYS4898 - Add to appliaction salutation if this is one of the first two applicants
        If (CSafeInt(xmlCustomerOrder.Text) <= 2) Then
           intAFFNumber = CSafeInt(vstrApplicationFFNumber)
           UpdateAppCorresSalutation vstrApplicationNumber, intAFFNumber
        End If
        'BMIDS00566 End
    End If
    
    ImportCustomers = blnOtherCustomers
        
    Set objCustomerBO = Nothing
    Set objICustomerRoleBO = Nothing
    Set objComboDO = Nothing
    Set xmlCustomerList = Nothing
    Set xmlCustomer = Nothing
    Set xmlCustRoleDoc = Nothing
    Set xmlTempResponseDoc = Nothing
    Set xmlTableElem = Nothing
    Set xmlElement = Nothing
    Set xmlRequestElem = Nothing
    Set xmlCustomerNo = Nothing
    Set xmlCustomerVersionNo = Nothing
    Set xmlCustomerRoleType = Nothing
    Set xmlCustomerOrder = Nothing
    Set xmlOtherSysCustNo = Nothing
    Set xmlNewCustNo = Nothing
    Set xmlNewCustVersionNo = Nothing
    Set xmlFindCustomerDoc = Nothing
    Set xmlFindCustomerNo = Nothing             ' PSC 03/10/03 BMIDS612
    Set xmlCustomerDoc = Nothing
    Set xmlCustRequestElem = Nothing
    Set xmlConvertedCustElem = Nothing
    Set xmlActionElem = Nothing
    Set xmlCustLockDoc = Nothing
    Set xmlLockCustNoElem = Nothing
    'BMIDS00004 MDC 21/05/2002 - Versions of Employment
    Set xmlTempRequest = Nothing
    Set xmlSearchNode = Nothing
    Set xmlCustomerNode = Nothing
    'BMIDS00004 MDC 21/05/2002 - End


    
    Exit Function

ImportCustomersVbErr:

    If m_objErrAssist.GetOmigaErrorNumber(Err.Number) = oeRecordNotFound Then
        Resume Next
    End If
    

    Set objCustomerBO = Nothing
    Set objICustomerRoleBO = Nothing
    Set objComboDO = Nothing
    Set xmlCustomerList = Nothing
    Set xmlCustomer = Nothing
    Set xmlCustRoleDoc = Nothing
    Set xmlTempResponseDoc = Nothing
    Set xmlTableElem = Nothing
    Set xmlElement = Nothing
    Set xmlRequestElem = Nothing
    Set xmlCustomerNo = Nothing
    Set xmlCustomerVersionNo = Nothing
    Set xmlCustomerRoleType = Nothing
    Set xmlCustomerOrder = Nothing
    Set xmlOtherSysCustNo = Nothing
    Set xmlNewCustNo = Nothing
    Set xmlNewCustVersionNo = Nothing
    Set xmlFindCustomerDoc = Nothing
    Set xmlCustomerDoc = Nothing
    Set xmlCustRequestElem = Nothing
    Set xmlConvertedCustElem = Nothing
    Set xmlActionElem = Nothing
    Set xmlCustLockDoc = Nothing
    Set xmlLockCustNoElem = Nothing
    
    m_objErrAssist.AddToErrSource cstrFunctionName
    
    Err.Raise Err.Number, Err.Source, Err.Description

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
    Dim strOmigaCustomerNumber As String    'BMIDS00701
    
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
    
    'BMIDS00701 Admin Customer may have Omiga3 customer number which should be used if available
    strOmigaCustomerNumber = Trim(m_objXmlAssist.GetNodeValue(xmlCustomerCopy, ".//OMIGACUSTOMERNUMBER", False, False))
    If Len(strOmigaCustomerNumber) > 0 Then
        Set xmlElement = xmlOut.createElement("OMIGACUSTOMERNUMBER")
        xmlElement.Text = strOmigaCustomerNumber
        xmlTableElem.appendChild xmlElement
    End If
    'BMIDS00701 End
    
    Set xmlElement = m_objXmlAssist.GetNode(xmlCustomerCopy, ".//OTHERSYSTEMTYPE")
    
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
        'BM0308 Dates only containing separator characters should be treated as blank
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
    
    'KW BMIDS745
    Set xmlElement = m_objXmlAssist.GetNode(xmlCustomerCopy, ".//NATIONALITY", False)
    
    If Not xmlElement Is Nothing Then
        xmlCustomerVersElem.appendChild xmlElement
    End If
    'KW End BMIDS745
    
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
    ' BMIDS864 End
    
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
    
    m_objErrAssist.AddToErrSource cstrFunctionName
    
    Err.Raise Err.Number, Err.Source, Err.Description

End Function

Private Function ImportPropertyInsuranceDetails(ByVal vxmlRequest As IXMLDOMElement, _
                                                ByVal vxmlProperty As IXMLDOMElement, _
                                                ByVal vxmlResponse As IXMLDOMElement) As String
' header ----------------------------------------------------------------------------------
' description:  Checks if the address already exists, if not creates the new address. Sets up
'               the Property Insurance Details
' pass:         vxmlRequest                     xml containing Request
'               vxmlProperty                    xml containing Property details
'               vxmlResponse                    xml to attach warnings to
' return:       ImportPropertyInsuranceDetails  Sequence number of the address
'------------------------------------------------------------------------------------------

On Error GoTo ImportPropertyInsuranceDetailsVbErr

    Const cstrFunctionName As String = "ImportPropertyInsuranceDetails"

    
    Dim objCustomerBO As CustomerBO
    Dim objComboDO As ComboDO
    Dim xmlRequestElem As IXMLDOMElement
    Dim xmlTableElem As IXMLDOMElement
    Dim xmlElement As IXMLDOMElement
    Dim xmlReponseDoc As FreeThreadedDOMDocument40
    Dim xmlCustAddressElem As IXMLDOMElement
    Dim xmlCustNo As IXMLDOMElement
    Dim xmlCustVersNo As IXMLDOMElement
    Dim xmlSeqNo As IXMLDOMElement
    Dim xmlCustomer As IXMLDOMElement
    Dim xmlCustomerAddressDoc As FreeThreadedDOMDocument40
    Dim xmlPropertyInsuranceDoc As FreeThreadedDOMDocument40
    
    Dim strResponse As String
    Dim strHouseNumber As String
    Dim strHouseName As String
    Dim strFlatNumber As String
    Dim strPostcode As String
    Dim strPattern As String
    Dim strPropertySeqNo As String
    Dim strAdditionalAddressValue As String
    
    ' Check if any of the customer addresses match the property address
    Set objCustomerBO = m_objContext.CreateInstance(gstrCUSTOMER_COMPONENT & ".CustomerBO")
    
    Set xmlCustomerAddressDoc = New FreeThreadedDOMDocument40
    xmlCustomerAddressDoc.validateOnParse = False
    xmlCustomerAddressDoc.setProperty "NewParser", True
    xmlCustomerAddressDoc.async = False
    
    Set xmlRequestElem = m_objXmlAssist.GetRequestNodeEx(vxmlRequest)
    xmlCustomerAddressDoc.appendChild xmlRequestElem
    
    Set xmlTableElem = xmlCustomerAddressDoc.createElement("CUSTOMERADDRESS")
    xmlRequestElem.appendChild xmlTableElem
    
    Set xmlElement = m_objXmlAssist.GetNode(vxmlRequest, ".//CUSTOMERNUMBER")
    Set xmlCustNo = xmlElement.cloneNode(True)
    xmlTableElem.appendChild xmlCustNo
    
    Set xmlElement = m_objXmlAssist.GetNode(vxmlRequest, ".//CUSTOMERVERSIONNUMBER")
    Set xmlCustVersNo = xmlElement.cloneNode(True)
    xmlTableElem.appendChild xmlCustVersNo
    
    strResponse = objCustomerBO.FindCustomerAddressList(xmlCustomerAddressDoc.xml)
    Set xmlReponseDoc = m_objXmlAssist.load(strResponse, TypeName(Me), cstrFunctionName)
    m_objErrAssist.CheckResponse xmlReponseDoc.documentElement, vxmlResponse, True
    
    strHouseNumber = m_objXmlAssist.GetNodeValue(vxmlProperty, ".//BUILDINGORHOUSENUMBER", False, False)
    strHouseName = m_objXmlAssist.GetNodeValue(vxmlProperty, ".//BUILDINGORHOUSENAME", False, False)
    strFlatNumber = m_objXmlAssist.GetNodeValue(vxmlProperty, ".//FLATNUMBER", False, False)
    strPostcode = m_objXmlAssist.GetNodeValue(vxmlProperty, ".//POSTCODE")
    
    If Len(strHouseNumber) = 0 And Len(strHouseName) = 0 Then
        m_objErrAssist.ThrowError TypeName(Me), _
                                  cstrFunctionName, _
                                  oeMissingElementValue, _
                                  "BUILDINGORHOUSNUMBER = " & strHouseNumber & _
                                  "; BUILDINGORHOUSNAME = " & strHouseName
    End If
    
    
    strPattern = ".//CUSTOMERADDRESS[ADDRESS[translate(POSTCODE, 'abcdefghijklmnopqrstuvwxyz', 'ABCDEFGHIJKLMNOPQRSTUVWXYZ') = '" & UCase$(strPostcode) & "'"
    
    If Len(strHouseNumber) > 0 Then
        strPattern = strPattern & " and translate(BUILDINGORHOUSENUMBER, 'abcdefghijklmnopqrstuvwxyz', 'ABCDEFGHIJKLMNOPQRSTUVWXYZ') = '" & UCase$(strHouseNumber) & "'"
    ElseIf Len(strHouseName) > 0 Then
        strPattern = strPattern & " and translate(BUILDINGORHOUSENAME, 'abcdefghijklmnopqrstuvwxyz', 'ABCDEFGHIJKLMNOPQRSTUVWXYZ') = '" & UCase$(strHouseName) & "'"
    End If
     
    If Len(strFlatNumber) > 0 Then
        strPattern = strPattern & " and translate(FLATNUMBER, 'abcdefghijklmnopqrstuvwxyz', 'ABCDEFGHIJKLMNOPQRSTUVWXYZ') = '" & UCase$(strFlatNumber) & "'"
    End If

    strPattern = strPattern & "]]"
    
    Set xmlCustAddressElem = xmlReponseDoc.selectSingleNode(strPattern)
    
    Set xmlReponseDoc = Nothing
    
    ' If the address exists get the sequence number from this else
    ' create a new address and get the sequence number from this
    If Not xmlCustAddressElem Is Nothing Then
        strPropertySeqNo = m_objXmlAssist.GetNodeValue(xmlCustAddressElem, _
                                                       ".//CUSTOMERADDRESSSEQUENCENUMBER")
    Else
        ' Get the combo value for an additional address
        Set objComboDO = m_objContext.CreateInstance(gstrBASE_COMPONENT & ".ComboDO")
        strAdditionalAddressValue = objComboDO.GetFirstComboValueId("CustomerAddressType", "A")

        Set xmlElement = xmlCustomerAddressDoc.createElement("ADDRESSTYPE")
        xmlElement.Text = strAdditionalAddressValue
        xmlTableElem.appendChild xmlElement
        
        Set xmlElement = xmlCustomerAddressDoc.createElement("ADDRESSGUID")
        xmlElement.Text = strAdditionalAddressValue
        xmlTableElem.appendChild xmlElement

        ' Create a new address
        Set xmlElement = m_objXmlAssist.GetNode(vxmlProperty, ".//PROPERTYADDRESS")
        xmlTableElem.appendChild xmlElement.cloneNode(True)
        m_objXmlAssist.ChangeNodeName xmlTableElem, "PROPERTYADDRESS", "ADDRESS"
        
        strResponse = objCustomerBO.CreateCustomerAddress(xmlCustomerAddressDoc.xml)
        Set xmlReponseDoc = m_objXmlAssist.load(strResponse, TypeName(Me), cstrFunctionName)
        m_objErrAssist.CheckResponse xmlReponseDoc.documentElement, vxmlResponse, True
        
        strPropertySeqNo = m_objXmlAssist.GetNodeValue(xmlReponseDoc, _
                                                       ".//CUSTOMERADDRESSSEQUENCENUMBER")
    End If
    
    Set xmlCustAddressElem = Nothing
    Set xmlCustomerAddressDoc = Nothing
        
    Set xmlPropertyInsuranceDoc = New FreeThreadedDOMDocument40
    xmlPropertyInsuranceDoc.validateOnParse = False
    xmlPropertyInsuranceDoc.setProperty "NewParser", True
    xmlPropertyInsuranceDoc.async = False
    Set xmlRequestElem = m_objXmlAssist.GetRequestNodeEx(vxmlRequest)
    xmlPropertyInsuranceDoc.appendChild xmlRequestElem

    Set xmlCustomer = xmlPropertyInsuranceDoc.createElement("CUSTOMER")
    xmlRequestElem.appendChild xmlCustomer
    
    Set xmlTableElem = xmlPropertyInsuranceDoc.createElement("CURRENTPROPERTY")
    xmlCustomer.appendChild xmlTableElem
    
    xmlTableElem.appendChild xmlCustNo
    xmlTableElem.appendChild xmlCustVersNo
    
    Set xmlSeqNo = xmlPropertyInsuranceDoc.createElement("CUSTOMERADDRESSSEQUENCENUMBER")
    xmlSeqNo.Text = strPropertySeqNo
    xmlTableElem.appendChild xmlSeqNo
    
    ' See if the property insurance details already exist
    strResponse = objCustomerBO.GetPropertyInsuranceDetails(xmlPropertyInsuranceDoc.xml)
    Set xmlReponseDoc = m_objXmlAssist.load(strResponse, TypeName(Me), cstrFunctionName)
    m_objErrAssist.CheckResponse xmlReponseDoc.documentElement, vxmlResponse, True
    Set xmlElement = m_objXmlAssist.GetNode(xmlReponseDoc, ".//CURRENTPROPERTY", False)
    
    ' Only create property insurance details if they don't exist
    If xmlElement Is Nothing Then
    
        Dim strValue As String
        
        strValue = m_objXmlAssist.GetNodeValue(vxmlProperty, ".//PURCHASEPRICE", False, False)
        Set xmlElement = xmlPropertyInsuranceDoc.createElement("ESTIMATEDCURRENTVALUE")
        xmlElement.Text = strValue
        xmlTableElem.appendChild xmlElement
        
        strValue = m_objXmlAssist.GetNodeValue(vxmlProperty, ".//VALUATIONDATE", False, False)
        Set xmlElement = xmlPropertyInsuranceDoc.createElement("LASTVALUATIONDATE")
        xmlElement.Text = strValue
        xmlTableElem.appendChild xmlElement
        
        strValue = m_objXmlAssist.GetNodeValue(vxmlProperty, ".//VALUATIONAMOUNT", False, False)
        Set xmlElement = xmlPropertyInsuranceDoc.createElement("LASTVALUATIONAMOUNT")
        xmlElement.Text = strValue
        xmlTableElem.appendChild xmlElement
        
        strValue = m_objXmlAssist.GetNodeValue(vxmlProperty, ".//VALUERID", False, False)
        Set xmlElement = xmlPropertyInsuranceDoc.createElement("LASTVALUERID")
        xmlElement.Text = strValue
        xmlTableElem.appendChild xmlElement
        
        strValue = m_objXmlAssist.GetNodeValue(vxmlProperty, ".//PROPERTYDESCRIPTION", False, False)
        Set xmlElement = xmlPropertyInsuranceDoc.createElement("DESCRIPTION")
        xmlElement.Text = strValue
        xmlTableElem.appendChild xmlElement
        
        strValue = m_objXmlAssist.GetNodeValue(vxmlProperty, ".//YEARBUILT", False, False)
        Set xmlElement = xmlPropertyInsuranceDoc.createElement("YEARBUILT")
        xmlElement.Text = strValue
        xmlTableElem.appendChild xmlElement
        
        strValue = m_objXmlAssist.GetNodeValue(vxmlProperty, ".//PROPERTYTENURE", False, False)
        Set xmlElement = xmlPropertyInsuranceDoc.createElement("TENURE")
        xmlElement.Text = strValue
        xmlTableElem.appendChild xmlElement
        
        strValue = m_objXmlAssist.GetNodeValue(vxmlProperty, ".//REINSTATEMENTAMOUNT", False, False)
        Set xmlElement = xmlPropertyInsuranceDoc.createElement("REINSTATEMENTAMOUNT")
        xmlElement.Text = strValue
        xmlTableElem.appendChild xmlElement
        
        Set xmlTableElem = xmlPropertyInsuranceDoc.createElement("HOMEINSURANCE")
        xmlCustomer.appendChild xmlTableElem
        xmlTableElem.appendChild xmlCustNo.cloneNode(True)
        xmlTableElem.appendChild xmlCustVersNo.cloneNode(True)
        xmlTableElem.appendChild xmlSeqNo.cloneNode(True)
        
        strValue = m_objXmlAssist.GetNodeValue(vxmlProperty, ".//BUILDINGSSUMINSURED", False, False)
        Set xmlElement = xmlPropertyInsuranceDoc.createElement("BUILDINGSSUMINSURED")
        xmlElement.Text = strValue
        xmlTableElem.appendChild xmlElement
        
        strValue = m_objXmlAssist.GetNodeValue(vxmlProperty, ".//HOMEINSURANCETYPE", False, False)
        Set xmlElement = xmlPropertyInsuranceDoc.createElement("HOMEINSURANCETYPE")
        xmlElement.Text = strValue
        xmlTableElem.appendChild xmlElement
        
        strResponse = objCustomerBO.CreatePropertyInsuranceDetails(xmlPropertyInsuranceDoc.xml)
        Set xmlReponseDoc = m_objXmlAssist.load(strResponse, TypeName(Me), cstrFunctionName)
        m_objErrAssist.CheckResponse xmlReponseDoc.documentElement, vxmlResponse, True
        
    End If
    
    ImportPropertyInsuranceDetails = strPropertySeqNo
    
    Set objCustomerBO = Nothing
    Set objComboDO = Nothing
    Set xmlRequestElem = Nothing
    Set xmlTableElem = Nothing
    Set xmlElement = Nothing
    Set xmlReponseDoc = Nothing
    Set xmlCustAddressElem = Nothing
    Set xmlCustNo = Nothing
    Set xmlCustVersNo = Nothing
    Set xmlSeqNo = Nothing
    Set xmlCustomer = Nothing
    Set xmlCustomerAddressDoc = Nothing
    Set xmlPropertyInsuranceDoc = Nothing


    
    Exit Function
                                                
ImportPropertyInsuranceDetailsVbErr:

    If m_objErrAssist.GetOmigaErrorNumber(Err.Number) = oeRecordNotFound Then
        Resume Next
    End If
    
    
    Set objCustomerBO = Nothing
    Set objComboDO = Nothing
    Set xmlRequestElem = Nothing
    Set xmlTableElem = Nothing
    Set xmlElement = Nothing
    Set xmlReponseDoc = Nothing
    Set xmlCustAddressElem = Nothing
    Set xmlCustNo = Nothing
    Set xmlCustVersNo = Nothing
    Set xmlSeqNo = Nothing
    Set xmlCustomer = Nothing
    Set xmlCustomerAddressDoc = Nothing
    Set xmlPropertyInsuranceDoc = Nothing
    
    m_objErrAssist.AddToErrSource cstrFunctionName

    Err.Raise Err.Number, Err.Source, Err.Description

End Function

Private Function ImportPropertyAddress(ByVal vxmlRequest As IXMLDOMElement, _
                                                ByVal vxmlProperty As IXMLDOMElement, _
                                                ByVal vxmlResponse As IXMLDOMElement) As String
' header ----------------------------------------------------------------------------------
' description:  Checks if the address already exists, if not creates the new address. Sets up
'               the Property Insurance Details
' pass:         vxmlRequest                     xml containing Request
'               vxmlProperty                    xml containing Property details
'               vxmlResponse                    xml to attach warnings to
' return:       ImportPropertyAddress           Address Guid of the address
' history:
' RF    06/10/03    BMIDS642     Fix problems with apostrophes within application / existing account data
'                                Improve error handling.
'------------------------------------------------------------------------------------------
On Error GoTo ImportPropertyAddressVbErr

    Const cstrFunctionName As String = "ImportPropertyAddress"

    'Dim objCustomerBO As CustomerBO
    'Dim objComboDO As ComboDO
    'Dim xmlRequestElem As IXMLDOMElement
    Dim xmlTableElem As IXMLDOMElement
    Dim xmlElement As IXMLDOMElement
    'Dim xmlResponseDoc As FreeThreadedDOMDocument40
    Dim xmlCustAddressElem As IXMLDOMElement
    'Dim xmlCustNo As IXMLDOMElement
    'Dim xmlCustVersNo As IXMLDOMElement
    'Dim xmlCustomerAddressDoc As FreeThreadedDOMDocument40
    'BMIDS00682 MDC 24/10/2002
    Dim xmlTempResponse As IXMLDOMNode
    Dim objIAddressBO As IAddressBO
    'BMIDS00682 MDC 24/10/2002 - End
    
    'Dim strResponse As String
    Dim strHouseNumber As String
    Dim strHouseName As String
    Dim strFlatNumber As String
    Dim strPostcode As String
    Dim strPattern As String
    'BMIDS00682 MDC 24/10/2002
    'Dim strPropertySeqNo As String
    Dim strPropertyAddressGuid As String
    'BMIDS00682 MDC 24/10/2002 - End
    'Dim strAdditionalAddressValue As String
    Dim strStreet   As String   'BMIDS00758
    
    'BMIDS918 GHun Use a cached CustomerAddress rather than retrieve the same thing for every account
    Dim xmlCustomerAddress As IXMLDOMNode
    Dim strPattern2         As String
    
    ' Check if any of the customer addresses match the property address
    'Set objCustomerBO = m_objContext.CreateInstance(gstrCUSTOMER_COMPONENT & ".CustomerBO")
    
    'Set xmlCustomerAddressDoc = New FreeThreadedDOMDocument40
    'xmlCustomerAddressDoc.async = False
    '
    'Set xmlRequestElem = m_objXmlAssist.GetRequestNodeEx(vxmlRequest)
    'xmlCustomerAddressDoc.appendChild xmlRequestElem
    '
    'Set xmlTableElem = xmlCustomerAddressDoc.createElement("CUSTOMERADDRESS")
    'xmlRequestElem.appendChild xmlTableElem
    '
    'Set xmlElement = m_objXmlAssist.GetNode(vxmlRequest, ".//CUSTOMERNUMBER")
    'Set xmlCustNo = xmlElement.cloneNode(True)
    'xmlTableElem.appendChild xmlCustNo
    '
    'Set xmlElement = m_objXmlAssist.GetNode(vxmlRequest, ".//CUSTOMERVERSIONNUMBER")
    'Set xmlCustVersNo = xmlElement.cloneNode(True)
    'xmlTableElem.appendChild xmlCustVersNo
    '
    'strResponse = objCustomerBO.FindCustomerAddressList(xmlCustomerAddressDoc.xml)
    'Set xmlResponseDoc = m_objXmlAssist.load(strResponse, TypeName(Me), cstrFunctionName)
    'm_objErrAssist.CheckResponse xmlResponseDoc.documentElement, vxmlResponse, True
    
    Set xmlCustomerAddress = vxmlRequest.selectSingleNode("CUSTOMERADDRESSLIST")
    'BMIDS918 End
    
    strHouseNumber = m_objXmlAssist.GetNodeValue(vxmlProperty, ".//BUILDINGORHOUSENUMBER", False, False)
    strHouseName = m_objXmlAssist.GetNodeValue(vxmlProperty, ".//BUILDINGORHOUSENAME", False, False)
    strFlatNumber = m_objXmlAssist.GetNodeValue(vxmlProperty, ".//FLATNUMBER", False, False)
    
    'BM0465 Postcodes can be blank
    'strPostcode = m_objXmlAssist.GetNodeValue(vxmlProperty, ".//POSTCODE")
    strPostcode = m_objXmlAssist.GetNodeValue(vxmlProperty, ".//POSTCODE", False, False)
    'BM0465
    
    'BMIDS00758
    strStreet = m_objXmlAssist.GetNodeValue(vxmlProperty, ".//STREET", False, False)
    
    'BMIDS00425 It is possible to have a valid address without a house name or number
    'If Len(strHouseNumber) = 0 And Len(strHouseName) = 0 Then
    '    m_objErrAssist.ThrowError TypeName(Me), _
    '                              cstrFunctionName, _
    '                              oeMissingElementValue, _
    '                              "BUILDINGORHOUSNUMBER = " & strHouseNumber & _
    '                              "; BUILDINGORHOUSNAME = " & strHouseName
    'End If
    'BMIDS00425
    
    ' RF 06/10/03 BMIDS642 Start
    'strPattern = ".//CUSTOMERADDRESS[ADDRESS[POSTCODE $ieq$ '" & strPostcode & "'"
    strPattern = "[translate(POSTCODE, 'abcdefghijklmnopqrstuvwxyz', 'ABCDEFGHIJKLMNOPQRSTUVWXYZ') = """ & UCase$(strPostcode) & """"
    
    If Len(strHouseNumber) > 0 Then
        'strPattern = strPattern & " and BUILDINGORHOUSENUMBER $ieq$ '" & strHouseNumber & "'"
        strPattern = strPattern & " and translate(BUILDINGORHOUSENUMBER, 'abcdefghijklmnopqrstuvwxyz', 'ABCDEFGHIJKLMNOPQRSTUVWXYZ') = """ & UCase$(strHouseNumber) & """"
    End If
    
    If Len(strHouseName) > 0 Then
        'strPattern = strPattern & " and BUILDINGORHOUSENAME $ieq$ '" & strHouseName & "'"
        strPattern = strPattern & " and translate(BUILDINGORHOUSENAME, 'abcdefghijklmnopqrstuvwxyz', 'ABCDEFGHIJKLMNOPQRSTUVWXYZ') = """ & UCase$(strHouseName) & """"
    End If
     
    If Len(strFlatNumber) > 0 Then
        'strPattern = strPattern & " and FLATNUMBER $ieq$ '" & strFlatNumber & "'"
        strPattern = strPattern & " and translate(FLATNUMBER, 'abcdefghijklmnopqrstuvwxyz', 'ABCDEFGHIJKLMNOPQRSTUVWXYZ') = """ & UCase$(strFlatNumber) & """"
    End If

    'BMIDS00758
    If Len(strStreet) > 0 Then
        'strPattern = strPattern & " and STREET $ieq$ '" & strStreet & "'"
        strPattern = strPattern & " and translate(STREET, 'abcdefghijklmnopqrstuvwxyz', 'ABCDEFGHIJKLMNOPQRSTUVWXYZ') = """ & UCase$(strStreet) & """"
    End If
    'BMIDS00758 End
    
    'BMIDS918 GHun
    strPattern2 = "//MORTGAGEACCOUNTLIST/MORTGAGEACCOUNT/PROPERTY/PROPERTYADDRESS" & strPattern & "]"
    strPattern = ".//CUSTOMERADDRESS[ADDRESS" & strPattern & "]]"
    'BMIDS918 End
    
    Set xmlCustAddressElem = xmlCustomerAddress.selectSingleNode(strPattern)
        
    ' If the address exists get the sequence number from this else
    ' create a new address and get the sequence number from this
    If Not xmlCustAddressElem Is Nothing Then
        'BMIDS00682 MDC 24/10/2002
'        strPropertySeqNo = m_objXmlAssist.GetNodeValue(xmlCustAddressElem, _
                                                       ".//CUSTOMERADDRESSSEQUENCENUMBER")
        strPropertyAddressGuid = m_objXmlAssist.GetNodeValue(xmlCustAddressElem, ".//ADDRESSGUID")
        'BMIDS00682 MDC 24/10/2002 - End
                                                       
    Else
        'BMIDS00682 MDC 24/10/2002 - Create Address only (not CustomerAddress) and return AddressGuid
'        ' Get the combo value for an additional address
'        Set objComboDO = m_objContext.CreateInstance(gstrBASE_COMPONENT & ".ComboDO")
'        strAdditionalAddressValue = objComboDO.GetFirstComboValueId("CustomerAddressType", "A")
'
'        Set xmlElement = xmlCustomerAddressDoc.createElement("ADDRESSTYPE")
'        xmlElement.Text = strAdditionalAddressValue
'        xmlTableElem.appendChild xmlElement
'
'        Set xmlElement = xmlCustomerAddressDoc.createElement("ADDRESSGUID")
'        xmlElement.Text = strAdditionalAddressValue
'        xmlTableElem.appendChild xmlElement
'
        ' Create a new address
        Set xmlElement = m_objXmlAssist.GetNode(vxmlProperty, ".//PROPERTYADDRESS")
        
        'BMIDS918 GHun Check if there is already an addressGUID for the property
        Set xmlTableElem = xmlElement.selectSingleNode("ADDRESSGUID")
        If xmlTableElem Is Nothing Then
        'BMIDS918 End
        
            m_objXmlAssist.ChangeNodeName xmlElement, "PROPERTYADDRESS", "ADDRESS"
    '        xmlTableElem.appendChild xmlElement.cloneNode(True)
    '        m_objXmlAssist.ChangeNodeName xmlTableElem, "PROPERTYADDRESS", "ADDRESS"
    '
    '        strResponse = objCustomerBO.CreateCustomerAddress(xmlCustomerAddressDoc.xml)
    '        Set xmlResponseDoc = m_objXmlAssist.load(strResponse, TypeName(Me), cstrFunctionName)
    '        m_objErrAssist.CheckResponse xmlResponseDoc.documentElement, vxmlResponse, True
    '
    '        strPropertySeqNo = m_objXmlAssist.GetNodeValue(xmlResponseDoc, _
    '                                                       ".//CUSTOMERADDRESSSEQUENCENUMBER")
            Set objIAddressBO = m_objContext.CreateInstance(App.Title & ".AddressBO")
            Set xmlTempResponse = objIAddressBO.CreateAddress(xmlElement)
            m_objErrAssist.CheckResponse xmlTempResponse, vxmlResponse, True
            'BMIDS918 GHun
            Set xmlTableElem = m_objXmlAssist.GetNode(xmlTempResponse, ".//ADDRESSGUID")
            'strPropertyAddressGuid = m_objXmlAssist.GetNodeValue(xmlTempResponse, ".//ADDRESSGUID")
            strPropertyAddressGuid = xmlTableElem.Text
            'BMIDS918 End
            'BMIDS00682 MDC 24/10/2002 - End
        
            'BMIDS918 GHun Append the new AddressGUID to other addresses that are the same
            For Each xmlElement In vxmlRequest.selectNodes(strPattern2)
                xmlElement.appendChild xmlTableElem.cloneNode(True)
            Next
        Else
            strPropertyAddressGuid = xmlTableElem.Text
        End If
        
        'BMIDS918 End
    
    End If
    
    'BMIDS00682 MDC 24/10/2002
    'ImportPropertyAddress = strPropertySeqNo
    ImportPropertyAddress = strPropertyAddressGuid
    'BMIDS00682 MDC 24/10/2002 - End
    
    'Set objCustomerBO = Nothing
    'Set objComboDO = Nothing
    'Set xmlRequestElem = Nothing
    Set xmlTableElem = Nothing
    Set xmlElement = Nothing
    'Set xmlResponseDoc = Nothing
    Set xmlCustAddressElem = Nothing
    'Set xmlCustNo = Nothing
    'Set xmlCustVersNo = Nothing
    'Set xmlCustomerAddressDoc = Nothing
    'BMIDS00682 MDC 24/10/2002
    Set xmlTempResponse = Nothing
    Set objIAddressBO = Nothing
    'BMIDS00682 MDC 24/10/2002 - End

    
    Exit Function
                                                
ImportPropertyAddressVbErr:

    If m_objErrAssist.GetOmigaErrorNumber(Err.Number) = oeRecordNotFound Then
        Resume Next
    End If
    
    'Set objCustomerBO = Nothing
    'Set objComboDO = Nothing
    'Set xmlRequestElem = Nothing
    Set xmlTableElem = Nothing
    Set xmlElement = Nothing
    'Set xmlResponseDoc = Nothing
    Set xmlCustAddressElem = Nothing
    'Set xmlCustNo = Nothing
    'Set xmlCustVersNo = Nothing
    'Set xmlCustomerAddressDoc = Nothing
    'BMIDS00682 MDC 24/10/2002
    Set xmlTempResponse = Nothing
    Set objIAddressBO = Nothing
    'BMIDS00682 MDC 24/10/2002 - End
    
    m_objErrAssist.AddToErrSource cstrFunctionName

    ' RF 06/10/03 BMIDS642
    If m_objErrAssist.IsSystemError = True Then
        m_objErrAssist.LogError TypeName(Me), cstrFunctionName, Err.Number, Err.Description
    End If
    
    Err.Raise Err.Number, Err.Source, Err.Description

End Function

Private Function ImportAccountAndLoans(ByVal vxmlRequest As IXMLDOMElement, _
                                       ByVal vxmlAccount As IXMLDOMElement, _
                                       ByVal vxmlMortgageAccountDoc As FreeThreadedDOMDocument40, _
                                       ByVal vxmlResponse As IXMLDOMElement, _
                                       ByVal vblnAdditionalInd As Boolean, _
                                       ByVal vstrAdditionalDetails As String, _
                                       ByVal vstrPropertyAddressGuid As String) As String   'BMIDS00682 MDC 24/10/2002
'                                       ByVal vstrPropertySeqNo As String) As String    'BMIDS00682 MDC 24/10/2002
' header ----------------------------------------------------------------------------------
' description:  Creates/updates the mortgage account and its repective loans
' pass:         vxmlRequest                     xml containing Request
'               vxmlAccount                     xml containing Account details
'               vxmlMortgageAccountDoc          xml holding current mortgage accounts for the
'                                               customer
'               vxmlResponse                    xml to attach warnings to
'               vblnAdditionalInd               Whether the account has additional customers
'               vstrAdditionalDetails           Details of the additional customers
'               vstrPropertyAddressGuid         The address guid of the account property
' return:       ImportAccountAndLoans           Guid of the created account
'------------------------------------------------------------------------------------------

On Error GoTo ImportAccountAndLoansVbErr

    Const cstrFunctionName As String = "ImportAccountAndLoans"
    
    Dim objMortgageLenderBO As MortgageLenderBO
    Dim objCustomerFinancialBO As CustomerFinancialBO
    Dim xmlResponseDoc As FreeThreadedDOMDocument40
    Dim xmlTableElem As IXMLDOMElement
    Dim xmlElement As IXMLDOMElement
    Dim xmlRequestElem As IXMLDOMElement
    Dim xmlDirectoryGuid As IXMLDOMElement
    Dim xmlAccountGuid As IXMLDOMElement
    Dim xmlLoan As IXMLDOMElement
    Dim xmlOmigaLoan As IXMLDOMElement
    Dim xmlLoanList As IXMLDOMNodeList
    Dim xmlOmigaLoanList As IXMLDOMNodeList
    Dim xmlMtgeLoanDoc As FreeThreadedDOMDocument40
    Dim xmlMtgeLenderDoc As FreeThreadedDOMDocument40
    Dim xmlOmigaMtgeAccount As IXMLDOMElement
    Dim xmlOmigaLoanDoc As FreeThreadedDOMDocument40
    
    ' PSC 01/02/2007 EP2_1113 - Start
    Dim objCrudBO As Object
    Dim xmlCrudRequest As IXMLDOMNode
    Dim xmlRoot As IXMLDOMNode
    Dim xmlCrudResponse As FreeThreadedDOMDocument40
    
    Dim strProdStartDate As String
    Dim strProductCode As String
    ' PSC 01/02/2007 EP2_1113 - End


    Dim strResponse As String
    Dim strAccountNumber As String
    Dim strPattern As String
    Dim strLoanAccountNo As String
            
    'BMIDS918 GHun Cache the Lender Directory GUID
    Dim strLenderCode As String
                 
    'Check if the DirectoryGUID is already in the XML
    Set xmlDirectoryGuid = vxmlAccount.selectSingleNode("DIRECTORYGUID")
    If xmlDirectoryGuid Is Nothing Then
    
        ' Get the directory guid of main lender
        Set xmlMtgeLenderDoc = New FreeThreadedDOMDocument40
        xmlMtgeLenderDoc.validateOnParse = False
        xmlMtgeLenderDoc.setProperty "NewParser", True
        xmlMtgeLenderDoc.async = False
        
        Set xmlRequestElem = m_objXmlAssist.GetRequestNodeEx(vxmlRequest)
        xmlMtgeLenderDoc.appendChild xmlRequestElem
        
        Set xmlTableElem = xmlMtgeLenderDoc.createElement("MORTGAGELENDER")
        xmlRequestElem.appendChild xmlTableElem
        
        Set xmlElement = m_objXmlAssist.GetNode(vxmlAccount, ".//LENDERCODE")
        strLenderCode = xmlElement.Text
        xmlTableElem.appendChild xmlElement.cloneNode(True)
        
        Set objMortgageLenderBO = _
                        m_objContext.CreateInstance(gstrORGANISATION_COMPONENT & ".MortgageLenderBO")
        strResponse = objMortgageLenderBO.GetMainLenderDirectoryGuid(xmlMtgeLenderDoc.xml)
        Set xmlResponseDoc = m_objXmlAssist.load(strResponse, TypeName(Me), cstrFunctionName)
        m_objErrAssist.CheckResponse xmlResponseDoc.documentElement, vxmlResponse, True
        
        Set xmlDirectoryGuid = m_objXmlAssist.GetNode(xmlResponseDoc, ".//DIRECTORYGUID")
        
        Set xmlMtgeLenderDoc = Nothing
        Set xmlResponseDoc = Nothing
        
        'Save the DirectoryGUID against all other MortgageAccounts that use the same lender,
        'so it does not have to be queried from the database again
        For Each xmlElement In vxmlRequest.selectNodes("//MORTGAGEACCOUNTLIST/MORTGAGEACCOUNT[LENDERCODE='" & strLenderCode & "']")
            xmlElement.appendChild xmlDirectoryGuid.cloneNode(True)
        Next
    End If
    'BMIDS918 End
    
    ' Determine whether to update the mortgage account or create it
    strAccountNumber = m_objXmlAssist.GetNodeValue(vxmlAccount, ".//ACCOUNTNUMBER")
    
    ' PSC 16/10/2002 BMIDS00575
    strPattern = ".//MORTGAGEACCOUNT[ACCOUNTNUMBER = '" & strAccountNumber & "']"
    
    Set xmlOmigaMtgeAccount = m_objXmlAssist.GetNode(vxmlMortgageAccountDoc, strPattern, False)
         
    If xmlOmigaMtgeAccount Is Nothing Then
        Set xmlAccountGuid = CreateMortgageAccount(vxmlRequest, _
                                                   vxmlAccount, _
                                                   xmlDirectoryGuid, _
                                                   vxmlResponse, _
                                                   vblnAdditionalInd, _
                                                   vstrAdditionalDetails, _
                                                   vstrPropertyAddressGuid) 'BMIDS00682 MDC 24/10/2002
                                                   'vstrPropertySeqNo)  'BMIDS00682 MDC 24/10/2002
    Else
        Set xmlAccountGuid = UpdateMortgageAccount(vxmlRequest, _
                                                   vxmlAccount, _
                                                   xmlOmigaMtgeAccount, _
                                                   xmlDirectoryGuid, _
                                                   vxmlResponse, _
                                                   vblnAdditionalInd, _
                                                   vstrAdditionalDetails, _
                                                   vstrPropertyAddressGuid) 'BMIDS00682 MDC 24/10/2002
                                                   'vstrPropertySeqNo)  'BMIDS00682 MDC 24/10/2002
    End If
    
            
    ' Set up a list of loans for this account if the account is found on Omiga. if they are not
    ' on the admin system delete them
    If Not xmlOmigaMtgeAccount Is Nothing Then
        Set xmlMtgeLoanDoc = New FreeThreadedDOMDocument40
        xmlMtgeLoanDoc.validateOnParse = False
        xmlMtgeLoanDoc.setProperty "NewParser", True
        xmlMtgeLoanDoc.async = False
        
        Set xmlRequestElem = m_objXmlAssist.GetRequestNodeEx(vxmlRequest)
        xmlMtgeLoanDoc.appendChild xmlRequestElem
        
        Set xmlTableElem = xmlMtgeLoanDoc.createElement("MORTGAGELOAN")
        xmlRequestElem.appendChild xmlTableElem
        xmlTableElem.appendChild xmlAccountGuid.cloneNode(True)
        
        Set objCustomerFinancialBO = _
                    m_objContext.CreateInstance(gstrCUSTOMER_FINANCIAL & ".CustomerFinancialBO")
        strResponse = objCustomerFinancialBO.FindMortgageLoanList(xmlMtgeLoanDoc.xml)
        Set xmlOmigaLoanDoc = m_objXmlAssist.load(strResponse, TypeName(Me), cstrFunctionName)
        m_objErrAssist.CheckResponse xmlOmigaLoanDoc.documentElement, vxmlResponse, True
    
        Set xmlOmigaLoanList = xmlOmigaLoanDoc.selectNodes(".//MORTGAGELOAN")
        
        ' Go through each Omiga loan. If it is not on the response from the admin system
        ' then delete it
        For Each xmlOmigaLoan In xmlOmigaLoanList
            strLoanAccountNo = m_objXmlAssist.GetNodeValue(xmlOmigaLoan, ".//LOANACCOUNTNUMBER")
            strPattern = ".//MORTGAGELOAN [LOANACCOUNTNUMBER = '" & strLoanAccountNo & "']"
            Set xmlLoan = m_objXmlAssist.GetNode(vxmlAccount, strPattern, False)
            
            If xmlLoan Is Nothing Then
                Set xmlElement = xmlTableElem
                Set xmlTableElem = xmlOmigaLoan.cloneNode(True)
                xmlRequestElem.replaceChild xmlTableElem, xmlElement
            
                strResponse = objCustomerFinancialBO.DeleteMortgageLoan(xmlMtgeLoanDoc.xml)
                Set xmlResponseDoc = m_objXmlAssist.load(strResponse, TypeName(Me), cstrFunctionName)
                m_objErrAssist.CheckResponse xmlResponseDoc.documentElement, vxmlResponse, True
                
                Set xmlResponseDoc = Nothing
            End If
        Next
    End If
    
    ' Set up list of loans off the admin system
    Set xmlLoanList = vxmlAccount.selectNodes(".//MORTGAGELOAN")
    
    If xmlLoanList.length > 0 Then
        Set objCrudBO = m_objContext.CreateInstance(gstrCRUD_COMPONENT & ".omCrudBO")
        Set xmlCrudRequest = m_objXmlAssist.GetRequestNodeEx(vxmlRequest)
        
        xmlSetAttributeValue xmlCrudRequest, "CRUD_OP", "READ"
        xmlSetAttributeValue xmlCrudRequest, "ENTITY_REF", "MORTGAGEPRODUCT"

        Set xmlRoot = vxmlMortgageAccountDoc.createElement("MORTGAGEPRODUCT")
        xmlCrudRequest.appendChild xmlRoot

    End If
    
    For Each xmlLoan In xmlLoanList
                    
        Set xmlOmigaLoan = Nothing
        
        ' Only see if this loan is already present if we have existing loans
        If Not xmlOmigaLoanDoc Is Nothing Then
            strLoanAccountNo = m_objXmlAssist.GetNodeValue(xmlLoan, ".//LOANACCOUNTNUMBER")
            strPattern = ".//MORTGAGELOAN [LOANACCOUNTNUMBER = '" & strLoanAccountNo & "']"

            Set xmlOmigaLoan = m_objXmlAssist.GetNode(xmlOmigaLoanDoc, strPattern, False)
        End If
        
        'BMIDS756 GHun
        xmlChangeNodeName xmlLoan, "PARTANDPARTINTONLYAMOUNT", "ORIGINALPARTANDPARTINTONLYAMT"
        'BMIDS756 End
        
        '**SR 15/01/2007 : EP2_858
        xmlChangeNodeName xmlLoan, "INCOMESTATUS", "ORIGINALINCOMESTATUS"  ' --??** discuss with RClay
        '-- Set the appropriate Redemption Status
        Dim strTypeOfApplication As String, strRedemptionStatus As String
        
        strTypeOfApplication = m_objXmlAssist.GetNodeValue(vxmlRequest, ".//TYPEOFAPPLICATION", True)
        ' PSC 16/02/2007 EP2_1314
        If IsValidationType("TypeOfMortgage", strTypeOfApplication, "F") And _
           IsValidationType("TypeOfMortgage", strTypeOfApplication, "M") Then
            strRedemptionStatus = GetFirstComboValueId("RedemptionStatus", "N")
        ElseIf IsValidationType("TypeOfMortgage", strTypeOfApplication, "PSW") Then
            ' PSC 30/03/2007 EP2_1636
            strRedemptionStatus = GetFirstComboValueId("RedemptionStatus", "PSR")
        ElseIf IsValidationType("TypeOfMortgage", strTypeOfApplication, "TOE") Then
            strRedemptionStatus = GetFirstComboValueId("RedemptionStatus", "TOE")
        ElseIf IsValidationType("TypeOfMortgage", strTypeOfApplication, "NP") Then
            strRedemptionStatus = GetFirstComboValueId("RedemptionStatus", "TBP")
        End If
        
        ' PSC 01/02/2007 EP2_1113 - Start
        Set xmlElement = xmlGetNode(xmlLoan, "REDEMPTIONSTATUS")
        
        If xmlElement Is Nothing Then
            Set xmlElement = vxmlMortgageAccountDoc.createElement("REDEMPTIONSTATUS")
            xmlLoan.appendChild xmlElement
        End If
        
        xmlElement.Text = strRedemptionStatus
        
        '**SR 15/01/2007 : EP2_858 - End
        
        If xmlOmigaLoan Is Nothing Then
            CreateMortgageLoan vxmlRequest, xmlLoan, xmlAccountGuid, vxmlResponse
        Else
            UpdateMortgageLoan vxmlRequest, xmlLoan, xmlOmigaLoan, xmlAccountGuid, vxmlResponse
        End If
    Next
    
    ImportAccountAndLoans = xmlAccountGuid.Text

    Set objMortgageLenderBO = Nothing
    Set objCustomerFinancialBO = Nothing
    Set xmlResponseDoc = Nothing
    Set xmlTableElem = Nothing
    Set xmlElement = Nothing
    Set xmlRequestElem = Nothing
    Set xmlDirectoryGuid = Nothing
    Set xmlAccountGuid = Nothing
    Set xmlLoan = Nothing
    Set xmlOmigaLoan = Nothing
    Set xmlLoanList = Nothing
    Set xmlOmigaLoanList = Nothing
    Set xmlMtgeLoanDoc = Nothing
    Set xmlMtgeLenderDoc = Nothing
    Set xmlOmigaMtgeAccount = Nothing
    Set xmlOmigaLoanDoc = Nothing
    
    ' PSC 01/02/2007 EP2_1113 - Start
    Set objCrudBO = Nothing
    Set xmlCrudRequest = Nothing
    Set xmlRoot = Nothing
    Set xmlCrudResponse = Nothing
    ' PSC 01/02/2007 EP2_1113 - End
    
    Exit Function
                                                
ImportAccountAndLoansVbErr:

    If m_objErrAssist.GetOmigaErrorNumber(Err.Number) = oeRecordNotFound Then
        Resume Next
    End If

    Set objMortgageLenderBO = Nothing
    Set objCustomerFinancialBO = Nothing
    Set xmlResponseDoc = Nothing
    Set xmlTableElem = Nothing
    Set xmlElement = Nothing
    Set xmlRequestElem = Nothing
    Set xmlDirectoryGuid = Nothing
    Set xmlAccountGuid = Nothing
    Set xmlLoan = Nothing
    Set xmlOmigaLoan = Nothing
    Set xmlLoanList = Nothing
    Set xmlOmigaLoanList = Nothing
    Set xmlMtgeLoanDoc = Nothing
    Set xmlMtgeLenderDoc = Nothing
    Set xmlOmigaMtgeAccount = Nothing
    Set xmlOmigaLoanDoc = Nothing
    
    ' PSC 01/02/2007 EP2_1113 - Start
    Set objCrudBO = Nothing
    Set xmlCrudRequest = Nothing
    Set xmlRoot = Nothing
    Set xmlCrudResponse = Nothing
    ' PSC 01/02/2007 EP2_1113 - End

    
    m_objErrAssist.AddToErrSource cstrFunctionName
    
    Err.Raise Err.Number, Err.Source, Err.Description

End Function

Private Function CreateMortgageAccount(ByVal vxmlRequest As IXMLDOMElement, _
                                       ByVal vxmlAccount As IXMLDOMElement, _
                                       ByVal vxmlDirectoryGuid As IXMLDOMElement, _
                                       ByVal vxmlResponse As IXMLDOMElement, _
                                       ByVal vblnAdditionalInd As Boolean, _
                                       ByVal vstrAdditionalDetails As String, _
                                       ByVal vstrPropertyAddressGuid As String) As IXMLDOMNode  'BMIDS00682 MDC 24/10/2002
                                       'ByVal vstrPropertySeqNo As String) As IXMLDOMNode    'BMIDS00682 MDC 24/10/2002
' header ----------------------------------------------------------------------------------
' description:  Creates the mortgage account with details from the administration system
' pass:         vxmlRequest                     xml containing Request
'               vxmlAccount                     xml containing Account details
'               vxmlResponse                    xml to attach warnings to
'               vblnAdditionalInd               Whether the account has additional customers
'               vstrAdditionalDetails           Details of the additional customers
'               vstrPropertyAddressGuid         The address guid of the account property
' return:       CreateMortgageAccount           Guid of the created account
'------------------------------------------------------------------------------------------

On Error GoTo CreateMortgageAccountVbErr

    Const cstrFunctionName As String = "CreateMortgageAccount"

    
    Dim xmlMtgeAccountElem As IXMLDOMElement
    Dim xmlResponseDoc As FreeThreadedDOMDocument40
    Dim xmlMtgeAccountDoc As FreeThreadedDOMDocument40
    Dim xmlRequestElem As IXMLDOMElement
    Dim xmlElement As IXMLDOMElement
    Dim xmlTableElem As IXMLDOMElement
    Dim objCustomerFinancialBO As CustomerFinancialBO

    'BMIDS00006 MDC 19/07/2002 - Customer Account Download
    Dim xmlGUIDNode As IXMLDOMNode
    Dim objCustomerBO As CustomerBO
    Dim xmlTempRequestElem As IXMLDOMElement
    Dim xmlAccountRelationshipNode As IXMLDOMNode
    Dim xmlCustomerListNode As IXMLDOMNode
    Dim xmlCustomerNode As IXMLDOMNode
    'BMIDS00006 MDC 19/07/2002 - End
    
    Dim xmlCustVersion As IXMLDOMNode       ' PSC 16/10/2002 BMIDS00575

    Dim strResponse As String
    Dim strCustomerNumber As String         ' PSC 16/10/2002 BMIDS00575
    Dim strPattern As String                ' PSC 16/10/2002 BMIDS00575
    
    Set objCustomerFinancialBO = _
                    m_objContext.CreateInstance(gstrCUSTOMER_FINANCIAL & ".CustomerFinancialBO")
    Set objCustomerBO = _
                    m_objContext.CreateInstance(gstrCUSTOMER_COMPONENT & ".CustomerBO")
   
    Set xmlMtgeAccountDoc = New FreeThreadedDOMDocument40
    xmlMtgeAccountDoc.validateOnParse = False
    xmlMtgeAccountDoc.setProperty "NewParser", True
    xmlMtgeAccountDoc.async = False
    
    Set xmlRequestElem = m_objXmlAssist.GetRequestNodeEx(vxmlRequest)
    xmlMtgeAccountDoc.appendChild xmlRequestElem
   
    Set xmlMtgeAccountElem = xmlMtgeAccountDoc.createElement("MORTGAGEACCOUNT")
    xmlRequestElem.appendChild xmlMtgeAccountElem
    
    'BMIDS00682 MDC 24/10/2002 - No need to find guid as it is now passed in
    'SR 16/01/2007 : EP2_858 - removed commented code. see v2 for removed code.

    Set xmlGUIDNode = xmlMtgeAccountDoc.createElement("ADDRESSGUID")
    xmlGUIDNode.Text = vstrPropertyAddressGuid
    xmlMtgeAccountElem.appendChild xmlGUIDNode

    'BMIDS00682 MDC 24/10/2002 - End
    
    Set xmlElement = xmlMtgeAccountDoc.createElement("ADDITIONALDETAILS")
    xmlElement.Text = vstrAdditionalDetails
    xmlMtgeAccountElem.appendChild xmlElement
    
    Set xmlElement = xmlMtgeAccountDoc.createElement("ADDITIONALINDICATOR")
    
    If vblnAdditionalInd = True Then
        xmlElement.Text = "1"
    Else
        xmlElement.Text = "0"
    End If
    
    xmlMtgeAccountElem.appendChild xmlElement
    
    Set xmlElement = m_objXmlAssist.GetNode(vxmlAccount, ".//SECONDCHARGEINDICATOR")
    xmlMtgeAccountElem.appendChild xmlElement.cloneNode(True)
        
    'BMIDS00006 MDC 19/07/2002 - Customer Account Download
    'Additional items moved from CurrentProperty and HomeInsurance
    Set xmlElement = xmlMtgeAccountDoc.createElement("DESCRIPTION")
    xmlElement.Text = m_objXmlAssist.GetNodeValue(vxmlAccount, ".//PROPERTYDESCRIPTION", False, False)
    xmlMtgeAccountElem.appendChild xmlElement.cloneNode(True)
        
    Set xmlElement = xmlMtgeAccountDoc.createElement("LASTVALUATIONAMOUNT")
    xmlElement.Text = m_objXmlAssist.GetNodeValue(vxmlAccount, ".//VALUATIONAMOUNT", False, False)
    xmlMtgeAccountElem.appendChild xmlElement.cloneNode(True)
        
    Set xmlElement = xmlMtgeAccountDoc.createElement("LASTVALUATIONDATE")
    xmlElement.Text = m_objXmlAssist.GetNodeValue(vxmlAccount, ".//VALUATIONDATE", False, False)
    xmlMtgeAccountElem.appendChild xmlElement.cloneNode(True)
        
    Set xmlElement = xmlMtgeAccountDoc.createElement("LASTVALUERID")
    xmlElement.Text = m_objXmlAssist.GetNodeValue(vxmlAccount, ".//VALUERID", False, False)
    xmlMtgeAccountElem.appendChild xmlElement.cloneNode(True)
        
    Set xmlElement = xmlMtgeAccountDoc.createElement("TENURE")
    xmlElement.Text = m_objXmlAssist.GetNodeValue(vxmlAccount, ".//PROPERTYTENURE", False, False)
    xmlMtgeAccountElem.appendChild xmlElement.cloneNode(True)
    
    Set xmlElement = m_objXmlAssist.GetNode(vxmlAccount, ".//YEARBUILT")
    xmlMtgeAccountElem.appendChild xmlElement.cloneNode(True)
    
    Set xmlElement = m_objXmlAssist.GetNode(vxmlAccount, ".//REINSTATEMENTAMOUNT")
    xmlMtgeAccountElem.appendChild xmlElement.cloneNode(True)
    
    Set xmlElement = m_objXmlAssist.GetNode(vxmlAccount, ".//BUILDINGSSUMINSURED")
    xmlMtgeAccountElem.appendChild xmlElement.cloneNode(True)
    
    Set xmlElement = m_objXmlAssist.GetNode(vxmlAccount, ".//HOMEINSURANCETYPE")
    xmlMtgeAccountElem.appendChild xmlElement.cloneNode(True)
    'BMIDS00006 MDC 19/07/2002 - End
    
    'GD BMIDS614 START
    'SR EP2_858 - add code to handle this value not being returned from Optimus
    Set xmlElement = xmlGetNode(vxmlAccount, ".//MONTHLYRENTALINCOME")
    If Not xmlElement Is Nothing Then
        xmlMtgeAccountElem.appendChild xmlElement.cloneNode(True)
    End If
    'GD BMIDS614 END
    
    'DRC BMIDS667 START
    Set xmlElement = xmlGetNode(vxmlAccount, ".//LASTVALUERNAME")  'SR EP2_858 - add code to handle this value not being returned from Optimus
    If Not xmlElement Is Nothing Then
        xmlMtgeAccountElem.appendChild xmlElement.cloneNode(True)
    End If
    ' Import Repossession flag
    Set xmlElement = xmlGetNode(vxmlAccount, ".//REPOSSESSIONFLAG") 'SR EP2_858 - add code to handle this value not being returned from Optimus
    If Not xmlElement Is Nothing Then
        xmlMtgeAccountElem.appendChild xmlElement.cloneNode(True)
    End If
    'DRC BMIDS667 END
    
    '*=[MC]START : BMIDS756
    'BMIDS756  Check that nodes exist before appending
    'BMIDS756 GHun The node may not exist, so pass 3rd parameter to GetNode as False
    'COLLATERAL ID
    Set xmlElement = m_objXmlAssist.GetNode(vxmlAccount, ".//COLLATERALID", False)
    If Not xmlElement Is Nothing Then
        xmlMtgeAccountElem.appendChild xmlElement.cloneNode(True)
    End If
    
    'TOTAL COLLATERAL BALANCE
    Set xmlElement = m_objXmlAssist.GetNode(vxmlAccount, ".//TOTALCOLLATERALBALANCE", False)
    If Not xmlElement Is Nothing Then
        xmlMtgeAccountElem.appendChild xmlElement.cloneNode(True)
    End If
    
    'PAYMENTDUEDATE
    Set xmlElement = m_objXmlAssist.GetNode(vxmlAccount, ".//PAYMENTDUEDATE", False)
    If Not xmlElement Is Nothing Then
        xmlMtgeAccountElem.appendChild xmlElement.cloneNode(True)
    End If
    
    'RMCFLAG
    'BMIDS756 should be REGULATEDMORTGAGEINDICIATOR not RMCFLAG
    Set xmlElement = m_objXmlAssist.GetNode(vxmlAccount, ".//REGULATEDMORTGAGEINDICIATOR", False)
    If Not xmlElement Is Nothing Then
        xmlChangeNodeName xmlElement, "REGULATEDMORTGAGEINDICIATOR", "RMCFLAG"
        xmlMtgeAccountElem.appendChild xmlElement.cloneNode(True)
    End If


    'DAILYMONTHLYANNUALINTEREST
    'BMIDS756 Should be INTERESTFREQUENCY not DAILYMONTHLYANNUALINTEREST
    Set xmlElement = m_objXmlAssist.GetNode(vxmlAccount, ".//INTERESTFREQUENCY", False)
    If Not xmlElement Is Nothing Then
        xmlChangeNodeName xmlElement, "INTERESTFREQUENCY", "DAILYMONTHLYANNUALINTEREST"

        xmlMtgeAccountElem.appendChild xmlElement.cloneNode(True)
    End If
    
    'DSSFLAG
    Set xmlElement = m_objXmlAssist.GetNode(vxmlAccount, ".//DSSFLAG", False)
    If Not xmlElement Is Nothing Then
        If (StrComp(xmlElement.Text, "DSS", vbTextCompare) = 0) Then
            xmlElement.Text = "1"
        Else
            xmlElement.Text = "0"
        End If

        xmlMtgeAccountElem.appendChild xmlElement.cloneNode(True)
    End If
    
    'BANKSORTCODE
    Set xmlElement = m_objXmlAssist.GetNode(vxmlAccount, ".//BANKSORTCODE", False)
    If Not xmlElement Is Nothing Then
        'BMIDS756 Reformat the sortcode if required
        Dim strBankSort As String
        Dim strNewBankSort As String
        Dim nLenSort As String
        strBankSort = xmlElement.Text
        If Len(strBankSort) = 6 Then
            nLenSort = InStr(1, strBankSort, "-", vbTextCompare)
            If nLenSort = 0 Then
                strNewBankSort = Left(strBankSort, 2) + "-" + Mid(strBankSort, 3, 2) + "-" + Right(strBankSort, 2)
                xmlElement.Text = strNewBankSort
            End If
        End If

        xmlMtgeAccountElem.appendChild xmlElement.cloneNode(True)
    End If
    
    'BANKACCOUNTNUMBER
    Set xmlElement = m_objXmlAssist.GetNode(vxmlAccount, ".//BANKACCOUNTNUMBER", False)
    If Not xmlElement Is Nothing Then
        xmlMtgeAccountElem.appendChild xmlElement.cloneNode(True)
    End If
    
    'BANKACCOUNTNAME
    Set xmlElement = m_objXmlAssist.GetNode(vxmlAccount, ".//BANKACCOUNTNAME", False)
    If Not xmlElement Is Nothing Then
        xmlMtgeAccountElem.appendChild xmlElement.cloneNode(True)
    End If
    
    'BUSINESS CHANNEL
    Set xmlElement = m_objXmlAssist.GetNode(vxmlAccount, ".//BUSINESSCHANNEL", False)
    If Not xmlElement Is Nothing Then
        xmlMtgeAccountElem.appendChild xmlElement.cloneNode(True)
    End If
    
    'BMIDS756 End
    '*=[MC]END : BMIDS756
    
    '**SR 15/01/2007 : EP2_858
    Set xmlElement = xmlGetNode(vxmlAccount, ".//ORIGINALNATUREOFLOAN")
    If Not xmlElement Is Nothing Then
        xmlMtgeAccountElem.appendChild xmlElement.cloneNode(True)
    End If
    
    Set xmlElement = xmlGetNode(vxmlAccount, ".//ORIGINALCREDITSCHEME")
    If Not xmlElement Is Nothing Then
        xmlMtgeAccountElem.appendChild xmlElement.cloneNode(True)
    End If
    
    Set xmlElement = xmlGetNode(vxmlAccount, ".//BANKNAME")
    If Not xmlElement Is Nothing Then
        xmlMtgeAccountElem.appendChild xmlElement.cloneNode(True)
    End If
    '**SR 15/01/2007 : EP2_858 - End
    
    ' PSC 25/01/2007 EP2_858 - Start
    Set xmlElement = xmlGetNode(vxmlAccount, ".//REDEMPTIONSTATUS")
    If Not xmlElement Is Nothing Then
        xmlMtgeAccountElem.appendChild xmlElement.cloneNode(True)
    End If
    
    Set xmlElement = xmlGetNode(vxmlAccount, ".//TOTALMONTHLYCOST")
    If Not xmlElement Is Nothing Then
        xmlMtgeAccountElem.appendChild xmlElement.cloneNode(True)
    End If

    Set xmlElement = xmlGetNode(vxmlAccount, ".//PREEMPTIONENDDATE")
    If Not xmlElement Is Nothing Then
        xmlMtgeAccountElem.appendChild xmlElement.cloneNode(True)
    End If
    ' PSC 25/01/2007 EP2_858 - End

    Set xmlTableElem = xmlMtgeAccountDoc.createElement("ACCOUNT")
    xmlMtgeAccountElem.appendChild xmlTableElem
            
    Set xmlElement = m_objXmlAssist.GetNode(vxmlAccount, ".//ACCOUNTNUMBER")
    xmlTableElem.appendChild xmlElement.cloneNode(True)
    
    xmlTableElem.appendChild vxmlDirectoryGuid
    
    Set xmlElement = xmlMtgeAccountDoc.createElement("DIRECTORYADDRESSINDICATOR")
    xmlElement.Text = "1"
    xmlTableElem.appendChild xmlElement
        
    Set xmlElement = xmlMtgeAccountDoc.createElement("IMPORTEDINDICATOR")
    xmlElement.Text = "1"
    xmlTableElem.appendChild xmlElement
    
    Set xmlTableElem = xmlMtgeAccountDoc.createElement("INDEMNITYINSURANCE")
    xmlMtgeAccountElem.appendChild xmlTableElem
    
    Set xmlElement = m_objXmlAssist.GetNode(vxmlAccount, ".//INDEMNITYCOMPANYNAME")
    xmlTableElem.appendChild xmlElement.cloneNode(True)
    
    Set xmlElement = m_objXmlAssist.GetNode(vxmlAccount, ".//INDEMNITYMORTGAGEAMOUNT")
    xmlTableElem.appendChild xmlElement.cloneNode(True)
    
    Set xmlElement = m_objXmlAssist.GetNode(vxmlAccount, ".//INDEMNITYAMOUNT")
    xmlTableElem.appendChild xmlElement.cloneNode(True)
    
    'BMIDS00006 MDC 19/07/2002 - Customer Account Download
    'Account Relationships
    Set xmlAccountRelationshipNode = xmlMtgeAccountDoc.createElement("ACCOUNTRELATIONSHIPLIST")
    'BMIDS00425
    
    ' PSC 16/10/2002 BMIDS00575 - Start
'    Set xmlCustomerListNode = m_objXmlAssist.GetNode(vxmlRequest, ".//CUSTOMERLIST", True)
    Set xmlCustomerListNode = m_objXmlAssist.GetNode(vxmlAccount, "CUSTOMERLIST", True)

'    For Each xmlCustomerNode In xmlCustomerListNode.childNodes
'        m_objXmlAssist.ChangeNodeName xmlCustomerNode, "CUSTOMER", "ACCOUNTRELATIONSHIP"
'        xmlAccountRelationshipNode.appendChild xmlCustomerNode.cloneNode(True)
'    Next
    
    
    For Each xmlCustomerNode In xmlCustomerListNode.childNodes
    
        ' Get the customer version number for this customer
        strCustomerNumber = m_objXmlAssist.GetNodeValue(xmlCustomerNode, "CUSTOMERNUMBER")
        strPattern = ".//CUSTOMERLIST/CUSTOMER[OTHERSYSTEMCUSTOMERNUMBER='" & strCustomerNumber & "']/CUSTOMERVERSIONNUMBER"
        Set xmlCustVersion = m_objXmlAssist.GetNode(vxmlRequest, strPattern)
        
        ' Copy the customer from the Account
        Set xmlElement = xmlCustomerNode.cloneNode(True)
        m_objXmlAssist.ChangeNodeName xmlElement, "CUSTOMER", "ACCOUNTRELATIONSHIP"
        m_objXmlAssist.ChangeNodeName xmlElement, "CUSTOMERNUMBER", "OTHERSYSTEMCUSTOMERNUMBER"
        m_objXmlAssist.ChangeNodeName xmlElement, "OMIGACUSTOMERNUMBER", "CUSTOMERNUMBER"
        
        ' Add in the customer version
        'BMIDS00682 MDC 25/10/2002 - Clone node before appending
        'xmlElement.appendChild xmlCustVersion
        xmlElement.appendChild xmlCustVersion.cloneNode(True)
        'BMIDS00682 MDC 25/10/2002 - End
        xmlAccountRelationshipNode.appendChild xmlElement
    Next
    ' PSC 16/10/2002 BMIDS00575 - End
    
    xmlMtgeAccountElem.appendChild xmlAccountRelationshipNode
    'BMIDS00425 End
    'BMIDS00006 MDC 19/07/2002 - End
    
    'BMIDS00508 MDC 24/09/2002 - Imported from BMids Admin System so must be a BMids account
    Set xmlElement = xmlMtgeAccountDoc.createElement("BMIDSACCOUNT")
    xmlElement.Text = "1"
    xmlMtgeAccountElem.appendChild xmlElement
    'BMIDS00508 MDC 24/09/2002 - End
    
    ' Create the mortgage account
    strResponse = objCustomerFinancialBO.CreateMortgageAccount(xmlMtgeAccountDoc.xml)
    Set xmlResponseDoc = m_objXmlAssist.load(strResponse, TypeName(Me), cstrFunctionName)
    m_objErrAssist.CheckResponse xmlResponseDoc.documentElement, vxmlResponse, True
        
    Set CreateMortgageAccount = m_objXmlAssist.GetNode(xmlResponseDoc, ".//ACCOUNTGUID")
    
    Set xmlMtgeAccountElem = Nothing
    Set xmlResponseDoc = Nothing
    Set xmlMtgeAccountDoc = Nothing
    Set xmlRequestElem = Nothing
    Set xmlElement = Nothing
    Set xmlTableElem = Nothing
    Set objCustomerFinancialBO = Nothing
    
    'BMIDS00006 MDC 19/07/2002 - Customer Account Download
    Set xmlGUIDNode = Nothing
    Set xmlTempRequestElem = Nothing
    Set objCustomerBO = Nothing
    Set xmlAccountRelationshipNode = Nothing
    Set xmlCustomerListNode = Nothing
    Set xmlCustomerNode = Nothing
    'BMIDS00006 MDC 19/07/2002 - End
    
    Set xmlCustVersion = Nothing       ' PSC 16/10/2002 BMIDS00575
    
    
    Exit Function
                                                
CreateMortgageAccountVbErr:

    Set xmlMtgeAccountElem = Nothing
    Set xmlResponseDoc = Nothing
    Set xmlMtgeAccountDoc = Nothing
    Set xmlRequestElem = Nothing
    Set xmlElement = Nothing
    Set xmlTableElem = Nothing
    Set objCustomerFinancialBO = Nothing
    
    
    m_objErrAssist.AddToErrSource cstrFunctionName
    
    Err.Raise Err.Number, Err.Source, Err.Description

End Function

Private Function UpdateMortgageAccount(ByVal vxmlRequest As IXMLDOMElement, _
                                       ByVal vxmlAccount As IXMLDOMElement, _
                                       ByVal vxmlOmigaAccount As IXMLDOMElement, _
                                       ByVal vxmlDirectoryGuid As IXMLDOMElement, _
                                       ByVal vxmlResponse As IXMLDOMElement, _
                                       ByVal vblnAdditionalInd As Boolean, _
                                       ByVal vstrAdditionalDetails As String, _
                                       ByVal vstrPropertyAddressGuid As String) As IXMLDOMElement   'BMIDS00682 MDC 24/10/2002
                                       'ByVal vstrPropertySeqNo As String) As IXMLDOMElement    'BMIDS00682 MDC 24/10/2002

' header ----------------------------------------------------------------------------------
' description:  Updates the mortgage account with the details from the administration system
' pass:         vxmlRequest                     xml containing Request
'               vxmlAccount                     xml containing Account details
'               vxmlOmigaAccount                xml holding equivalent mortgage account in Omiga
'               vxmlResponse                    xml to attach warnings to
'               vblnAdditionalInd               Whether the account has additional customers
'               vstrAdditionalDetails           Details of the additional customers
'               vstrPropertyAddressGuid         The Address Guid of the account property
' return:       UpdateMortgageAccount           Guid of the updated account
'------------------------------------------------------------------------------------------

On Error GoTo UpdateMortgageAccountVbErr

    Const cstrFunctionName As String = "UpdateMortgageAccount"

    
    Dim xmlMtgeAccountDoc As FreeThreadedDOMDocument40
    Dim xmlResponseDoc As FreeThreadedDOMDocument40
    Dim xmlRequestElem As IXMLDOMElement
    Dim xmlElement As IXMLDOMElement
    Dim xmlOmigaElement As IXMLDOMElement
    Dim xmlOmigaMtgeAccount As IXMLDOMElement
    Dim xmlAccount As IXMLDOMElement
    Dim xmlIndemnityInsurance As IXMLDOMElement
    Dim xmlTableElem As IXMLDOMElement
    Dim objCustomerFinancialBO As CustomerFinancialBO
    
    'BMIDS00006 MDC 22/07/2002 - Customer Account Download
    Dim xmlGUIDNode As IXMLDOMNode
    'BMIDS00682 MDC 24/10/2002 - No longer required
    'Dim objCustomerBO As CustomerBO
    'Dim xmlTempRequestElem As IXMLDOMElement
    'BMIDS00682 MDC 24/10/2002 - End
    Dim xmlAccountRelationshipNode As IXMLDOMNode
    Dim xmlCustomerListNode As IXMLDOMNode
    Dim xmlCustomerNode As IXMLDOMNode
    'BMIDS00006 MDC 22/07/2002 - End
    
    Dim xmlCustVersion As IXMLDOMNode       ' PSC 16/10/2002 BMIDS00575
    
    Dim strResponse As String
    Dim strPattern As String
    Dim strValue As String                  ' PSC 16/10/2002 BMIDS00575
    Dim strCustomerNumber As String         ' PSC 16/10/2002 BMIDS00575
    
    Set objCustomerFinancialBO = _
                    m_objContext.CreateInstance(gstrCUSTOMER_FINANCIAL & ".CustomerFinancialBO")
    'BMIDS00682 MDC 24/10/2002 - No longer required
    'BMIDS00006 MDC 22/07/2002 - Customer Account Download
    'Set objCustomerBO = _
                    m_objContext.CreateInstance(gstrCUSTOMER_COMPONENT & ".CustomerBO")
    'BMIDS00006 MDC 22/07/2002 - End
    'BMIDS00682 MDC 24/10/2002 - End
   
    Set xmlOmigaMtgeAccount = vxmlOmigaAccount.cloneNode(True)
    
    Set xmlMtgeAccountDoc = New FreeThreadedDOMDocument40
    xmlMtgeAccountDoc.validateOnParse = False
    xmlMtgeAccountDoc.setProperty "NewParser", True
    xmlMtgeAccountDoc.async = False
    
    Set xmlRequestElem = m_objXmlAssist.GetRequestNodeEx(vxmlRequest)
    xmlMtgeAccountDoc.appendChild xmlRequestElem
       
    'BMIDS00682 MDC 24/10/2002 - No need to find AddressGuid as now it is passed in
    'SR 16/01/2007 : EP2_858 - removed commented code for the above defect. see V2 for the removed code
    
    Set xmlGUIDNode = xmlMtgeAccountDoc.createElement("ADDRESSGUID")
    xmlGUIDNode.Text = vstrPropertyAddressGuid
    
    Set xmlElement = m_objXmlAssist.GetNode(xmlOmigaMtgeAccount, "ADDRESSGUID", False)
    If xmlElement Is Nothing Then
        xmlOmigaMtgeAccount.appendChild xmlGUIDNode
    Else
        xmlOmigaMtgeAccount.replaceChild xmlGUIDNode, xmlElement
    End If
    'BMIDS00682 MDC 24/10/2002 - End
    
    'PSC 16/10/2002 BMIDS00575 - Start
    'SR 16/01/2007 : EP2_858 - removed commented code for the above defect. see V2 for the removed code

    strPattern = ".//SECONDCHARGEINDICATOR"
    Set xmlOmigaElement = m_objXmlAssist.GetNode(xmlOmigaMtgeAccount, strPattern)
    Set xmlElement = m_objXmlAssist.GetNode(vxmlAccount, strPattern)
    Set xmlElement = xmlElement.cloneNode(True)
    xmlOmigaMtgeAccount.replaceChild xmlElement, xmlOmigaElement
       
    strValue = m_objXmlAssist.GetNodeValue(vxmlAccount, ".//PROPERTYDESCRIPTION", False, False)
    
    If Len(Trim$(strValue)) > 0 Then
        Set xmlOmigaElement = xmlMtgeAccountDoc.createElement("DESCRIPTION")
        xmlOmigaElement.Text = strValue
        xmlOmigaMtgeAccount.appendChild xmlOmigaElement
    End If
    
    strValue = m_objXmlAssist.GetNodeValue(vxmlAccount, ".//VALUATIONAMOUNT", False, False)
    
    If Len(Trim$(strValue)) > 0 Then
        Set xmlOmigaElement = xmlMtgeAccountDoc.createElement("LASTVALUATIONAMOUNT")
        xmlOmigaElement.Text = strValue
        xmlOmigaMtgeAccount.appendChild xmlOmigaElement
    End If
    
    strValue = m_objXmlAssist.GetNodeValue(vxmlAccount, ".//VALUATIONDATE", False, False)
    
    If Len(Trim$(strValue)) > 0 Then
        Set xmlOmigaElement = xmlMtgeAccountDoc.createElement("LASTVALUATIONDATE")
        xmlOmigaElement.Text = strValue
        xmlOmigaMtgeAccount.appendChild xmlOmigaElement
    End If
    
    'GD BMIDS614 START
    strValue = m_objXmlAssist.GetNodeValue(vxmlAccount, ".//MONTHLYRENTALINCOME", False, False)
    
    If Len(Trim$(strValue)) > 0 Then
        Set xmlOmigaElement = xmlMtgeAccountDoc.createElement("MONTHLYRENTALINCOME")
        xmlOmigaElement.Text = strValue
        xmlOmigaMtgeAccount.appendChild xmlOmigaElement
    End If
    'GD BMIDS614 END
    
   'DRC BMIDS667 START
    strValue = m_objXmlAssist.GetNodeValue(vxmlAccount, ".//LASTVALUERNAME", False, False)
    
    If Len(Trim$(strValue)) > 0 Then
        Set xmlOmigaElement = xmlMtgeAccountDoc.createElement("LASTVALUERNAME")
        xmlOmigaElement.Text = strValue
        xmlOmigaMtgeAccount.appendChild xmlOmigaElement
    End If
' Import Repossession flag
    strValue = m_objXmlAssist.GetNodeValue(vxmlAccount, ".//REPOSSESSIONFLAG", False, False)

    If Len(Trim$(strValue)) > 0 Then
        Set xmlOmigaElement = xmlMtgeAccountDoc.createElement("REPOSSESSIONFLAG")
        xmlOmigaElement.Text = strValue
        xmlOmigaMtgeAccount.appendChild xmlOmigaElement
    End If
    
       
    'DRC BMIDS667 END
    strValue = m_objXmlAssist.GetNodeValue(vxmlAccount, ".//VALUERID", False, False)
    
    If Len(Trim$(strValue)) > 0 Then
        Set xmlOmigaElement = xmlMtgeAccountDoc.createElement("LASTVALUERID")
        xmlOmigaElement.Text = strValue
        xmlOmigaMtgeAccount.appendChild xmlOmigaElement
    End If
        
    strValue = m_objXmlAssist.GetNodeValue(vxmlAccount, ".//PROPERTYTENURE", False, False)
    
    If Len(Trim$(strValue)) > 0 Then
        Set xmlOmigaElement = xmlMtgeAccountDoc.createElement("TENURE")
        xmlOmigaElement.Text = strValue
        xmlOmigaMtgeAccount.appendChild xmlOmigaElement
    End If
       
    strValue = m_objXmlAssist.GetNodeValue(vxmlAccount, ".//YEARBUILT", False, False)
    
    If Len(Trim$(strValue)) > 0 Then
        Set xmlOmigaElement = xmlMtgeAccountDoc.createElement("YEARBUILT")
        xmlOmigaElement.Text = strValue
        xmlOmigaMtgeAccount.appendChild xmlOmigaElement
    End If
    
    strValue = m_objXmlAssist.GetNodeValue(vxmlAccount, ".//REINSTATEMENTAMOUNT", False, False)
    
    If Len(Trim$(strValue)) > 0 Then
        Set xmlOmigaElement = xmlMtgeAccountDoc.createElement("REINSTATEMENTAMOUNT")
        xmlOmigaElement.Text = strValue
        xmlOmigaMtgeAccount.appendChild xmlOmigaElement
    End If
    
    strValue = m_objXmlAssist.GetNodeValue(vxmlAccount, ".//BUILDINGSSUMINSURED", False, False)
    
    If Len(Trim$(strValue)) > 0 Then
        Set xmlOmigaElement = xmlMtgeAccountDoc.createElement("BUILDINGSSUMINSURED")
        xmlOmigaElement.Text = strValue
        xmlOmigaMtgeAccount.appendChild xmlOmigaElement
    End If
    
    strValue = m_objXmlAssist.GetNodeValue(vxmlAccount, ".//HOMEINSURANCETYPE", False, False)
    
    If Len(Trim$(strValue)) > 0 Then
        Set xmlOmigaElement = xmlMtgeAccountDoc.createElement("HOMEINSURANCETYPE")
        xmlOmigaElement.Text = strValue
        xmlOmigaMtgeAccount.appendChild xmlOmigaElement
    End If
      
     '*=[MC]START : BMIDS756
    
    'COLLATERAL ID
    'Set xmlElement = m_objXmlAssist.GetNode(vxmlAccount, ".//COLLATERALID")
    'xmlMtgeAccountElem.appendChild xmlElement.cloneNode(True)
    
    strValue = m_objXmlAssist.GetNodeValue(vxmlAccount, ".//COLLATERALID", False, False)
    'If Len(Trim$(strValue)) > 0 Then
    Set xmlOmigaElement = xmlMtgeAccountDoc.createElement("COLLATERALID")
    xmlOmigaElement.Text = strValue
    xmlOmigaMtgeAccount.appendChild xmlOmigaElement
    'End If
    
    'TOTAL COLLATERAL BALANCE
    'Set xmlElement = m_objXmlAssist.GetNode(vxmlAccount, ".//TOTALCOLLATERALBALANCE")
    'xmlMtgeAccountElem.appendChild xmlElement.cloneNode(True)
    
    strValue = m_objXmlAssist.GetNodeValue(vxmlAccount, ".//TOTALCOLLATERALBALANCE", False, False)

    Set xmlOmigaElement = xmlMtgeAccountDoc.createElement("TOTALCOLLATERALBALANCE")
    xmlOmigaElement.Text = strValue
    xmlOmigaMtgeAccount.appendChild xmlOmigaElement

    
    
    'PAYMENTDUEDATE
    'Set xmlElement = m_objXmlAssist.GetNode(vxmlAccount, ".//PAYMENTDUEDATE")
    'xmlMtgeAccountElem.appendChild xmlElement.cloneNode(True)
    
    strValue = m_objXmlAssist.GetNodeValue(vxmlAccount, ".//PAYMENTDUEDATE", False, False)

    Set xmlOmigaElement = xmlMtgeAccountDoc.createElement("PAYMENTDUEDATE")
    xmlOmigaElement.Text = strValue
    xmlOmigaMtgeAccount.appendChild xmlOmigaElement

    'RMCFLAG
    'Set xmlElement = m_objXmlAssist.GetNode(vxmlAccount, ".//RMCFLAG")
    'xmlMtgeAccountElem.appendChild xmlElement.cloneNode(True)
    
    'BMIDS756 should be REGULATEDMORTGAGEINDICIATOR not RMCFLAG
    strValue = m_objXmlAssist.GetNodeValue(vxmlAccount, ".//REGULATEDMORTGAGEINDICIATOR", False, False)

    Set xmlOmigaElement = xmlMtgeAccountDoc.createElement("RMCFLAG")
    xmlOmigaElement.Text = strValue

    xmlOmigaMtgeAccount.appendChild xmlOmigaElement

    
    'DAILYMONTHLYANNUALINTEREST
    'Set xmlElement = m_objXmlAssist.GetNode(vxmlAccount, ".//DAILYMONTHLYANNUALINTEREST")
    'xmlMtgeAccountElem.appendChild xmlElement.cloneNode(True)
    
    'BMIDS756 Should be INTERESTFREQUENCY not DAILYMONTHLYANNUALINTEREST
    strValue = m_objXmlAssist.GetNodeValue(vxmlAccount, ".//INTERESTFREQUENCY", False, False)

    Set xmlOmigaElement = xmlMtgeAccountDoc.createElement("DAILYMONTHLYANNUALINTEREST")
    xmlOmigaElement.Text = strValue

    xmlOmigaMtgeAccount.appendChild xmlOmigaElement

    'DSSFLAG
    'Set xmlElement = m_objXmlAssist.GetNode(vxmlAccount, ".//DSSFLAG")
    'xmlMtgeAccountElem.appendChild xmlElement.cloneNode(True)
    
    strValue = m_objXmlAssist.GetNodeValue(vxmlAccount, ".//DSSFLAG", False, False)

    Set xmlOmigaElement = xmlMtgeAccountDoc.createElement("DSSFLAG")
    '
    If (StrComp(strValue, "DSS", vbTextCompare) = 0) Then
        xmlOmigaElement.Text = "1"
    Else
        xmlOmigaElement.Text = "0"
    End If
    xmlOmigaMtgeAccount.appendChild xmlOmigaElement

    
    'BANKSORTCODE
    'Set xmlElement = m_objXmlAssist.GetNode(vxmlAccount, ".//BANKSORTCODE")
    'xmlMtgeAccountElem.appendChild xmlElement.cloneNode(True)
    
    strValue = m_objXmlAssist.GetNodeValue(vxmlAccount, ".//BANKSORTCODE", False, False)
    'BMIDS756 Reformat the sortcode if required
    Dim strNewBankSort As String
    Dim nLenSort As String
    If Len(strValue) = 6 Then
        nLenSort = InStr(1, strValue, "-", vbTextCompare)
        If nLenSort = 0 Then
            strNewBankSort = Left(strValue, 2) + "-" + Mid(strValue, 3, 2) + "-" + Right(strValue, 2)
            strValue = strNewBankSort
        End If
    End If

    Set xmlOmigaElement = xmlMtgeAccountDoc.createElement("BANKSORTCODE")
    xmlOmigaElement.Text = strValue
    xmlOmigaMtgeAccount.appendChild xmlOmigaElement

    
    'BANKACCOUNTNUMBER
    'Set xmlElement = m_objXmlAssist.GetNode(vxmlAccount, ".//BANKACCOUNTNUMBER")
    'xmlMtgeAccountElem.appendChild xmlElement.cloneNode(True)
    strValue = m_objXmlAssist.GetNodeValue(vxmlAccount, ".//BANKACCOUNTNUMBER", False, False)

    Set xmlOmigaElement = xmlMtgeAccountDoc.createElement("BANKACCOUNTNUMBER")
    xmlOmigaElement.Text = strValue
    xmlOmigaMtgeAccount.appendChild xmlOmigaElement

    
    'BANKACCOUNTNAME
    'Set xmlElement = m_objXmlAssist.GetNode(vxmlAccount, ".//BANKACCOUNTNAME")
    'xmlMtgeAccountElem.appendChild xmlElement.cloneNode(True)
    
    strValue = m_objXmlAssist.GetNodeValue(vxmlAccount, ".//BANKACCOUNTNAME", False, False)

    Set xmlOmigaElement = xmlMtgeAccountDoc.createElement("BANKACCOUNTNAME")
    xmlOmigaElement.Text = strValue
    xmlOmigaMtgeAccount.appendChild xmlOmigaElement

    
    'BUSINESS CHANNEL
    'Set xmlElement = m_objXmlAssist.GetNode(vxmlAccount, ".//BUSINESSCHANNEL")
    'xmlMtgeAccountElem.appendChild xmlElement.cloneNode(True)
    
    strValue = m_objXmlAssist.GetNodeValue(vxmlAccount, ".//BUSINESSCHANNEL", False, False)

    Set xmlOmigaElement = xmlMtgeAccountDoc.createElement("BUSINESSCHANNEL")
    xmlOmigaElement.Text = strValue
    xmlOmigaMtgeAccount.appendChild xmlOmigaElement
     '*=[MC]END : BMIDS756
    
    '**SR 15/01/2007 : EP2_858
    strValue = m_objXmlAssist.GetNodeValue(vxmlAccount, ".//ORIGINALNATUREOFLOAN", False, False)
    Set xmlOmigaElement = xmlMtgeAccountDoc.createElement("ORIGINALNATUREOFLOAN")
    xmlOmigaElement.Text = strValue
    xmlOmigaMtgeAccount.appendChild xmlOmigaElement
    
    strValue = m_objXmlAssist.GetNodeValue(vxmlAccount, ".//ORIGINALCREDITSCHEME", False, False)
    Set xmlOmigaElement = xmlMtgeAccountDoc.createElement("ORIGINALCREDITSCHEME")
    xmlOmigaElement.Text = strValue
    xmlOmigaMtgeAccount.appendChild xmlOmigaElement
    '**SR 15/01/2007 : EP2_858 - End
    
    ' PSC 25/01/2007 EP2_858 - Start
    strValue = m_objXmlAssist.GetNodeValue(vxmlAccount, ".//REDEMPTIONSTATUS", False, False)
    Set xmlOmigaElement = xmlMtgeAccountDoc.createElement("REDEMPTIONSTATUS")
    xmlOmigaElement.Text = strValue
    xmlOmigaMtgeAccount.appendChild xmlOmigaElement
    
    strValue = m_objXmlAssist.GetNodeValue(vxmlAccount, ".//TOTALMONTHLYCOST", False, False)
    Set xmlOmigaElement = xmlMtgeAccountDoc.createElement("TOTALMONTHLYCOST")
    xmlOmigaElement.Text = strValue
    xmlOmigaMtgeAccount.appendChild xmlOmigaElement
    
    strValue = m_objXmlAssist.GetNodeValue(vxmlAccount, ".//PREEMPTIONENDDATE", False, False)
    Set xmlOmigaElement = xmlMtgeAccountDoc.createElement("PREEMPTIONENDDATE")
    xmlOmigaElement.Text = strValue
    xmlOmigaMtgeAccount.appendChild xmlOmigaElement
    ' PSC 25/01/2007 EP2_858 - End
    
      
    'Account ---------------------------------------------------------------------------
'    Set xmlAccount = m_objXmlAssist.GetNode(xmlOmigaMtgeAccount, ".//ACCOUNT")
    Set xmlGUIDNode = m_objXmlAssist.GetNode(vxmlOmigaAccount, "ACCOUNTGUID")

    Set xmlAccount = xmlMtgeAccountDoc.createElement("ACCOUNT")
    xmlOmigaMtgeAccount.appendChild xmlAccount
    xmlAccount.appendChild xmlGUIDNode.cloneNode(True)

    strPattern = ".//ACCOUNTNUMBER"
'    Set xmlOmigaElement = m_objXmlAssist.GetNode(xmlAccount, strPattern)
'    Set xmlElement = m_objXmlAssist.GetNode(vxmlAccount, strPattern)
'    Set xmlElement = xmlElement.cloneNode(True)
'    xmlAccount.replaceChild xmlElement, xmlOmigaElement
     
    strValue = m_objXmlAssist.GetNodeValue(vxmlAccount, strPattern, False, False)
    
    If Len(Trim$(strValue)) > 0 Then
        Set xmlOmigaElement = xmlMtgeAccountDoc.createElement("ACCOUNTNUMBER")
        xmlAccount.appendChild xmlOmigaElement
        xmlOmigaElement.Text = strValue
    End If
        
'    strPattern = ".//DIRECTORYGUID"
'    Set xmlOmigaElement = m_objXmlAssist.GetNode(xmlAccount, strPattern)
'    xmlAccount.replaceChild vxmlDirectoryGuid, xmlOmigaElement
    
    strValue = vxmlDirectoryGuid.Text
    
    If Len(Trim$(strValue)) > 0 Then
        Set xmlOmigaElement = xmlMtgeAccountDoc.createElement("DIRECTORYGUID")
        xmlAccount.appendChild xmlOmigaElement
        xmlOmigaElement.Text = strValue
    End If
    
'    Set xmlElement = xmlOmigaElement.ownerDocument.createElement("DIRECTORYADDRESSINDICATOR")
'    xmlElement.Text = "1"
'    xmlAccount.appendChild xmlElement

    Set xmlOmigaElement = xmlMtgeAccountDoc.createElement("DIRECTORYADDRESSINDICATOR")
    xmlAccount.appendChild xmlOmigaElement
    xmlOmigaElement.Text = "1"
    
'    strPattern = ".//IMPORTEDINDICATOR"
'    Set xmlOmigaElement = m_objXmlAssist.GetNode(xmlAccount, strPattern)
    Set xmlOmigaElement = xmlMtgeAccountDoc.createElement("IMPORTEDINDICATOR")
    xmlAccount.appendChild xmlOmigaElement
    xmlOmigaElement.Text = "1"
    
    
    'Indemnity Insurance ---------------------------------------------------------------
'    Set xmlIndemnityInsurance = m_objXmlAssist.GetNode(xmlOmigaMtgeAccount, ".//INDEMNITYINSURANCE")
    Set xmlIndemnityInsurance = xmlMtgeAccountDoc.createElement("INDEMNITYINSURANCE")
    xmlOmigaMtgeAccount.appendChild xmlIndemnityInsurance
    xmlIndemnityInsurance.appendChild xmlGUIDNode.cloneNode(True)
    
    strPattern = ".//INDEMNITYCOMPANYNAME"
'    Set xmlOmigaElement = m_objXmlAssist.GetNode(xmlIndemnityInsurance, strPattern)
'    Set xmlElement = m_objXmlAssist.GetNode(vxmlAccount, strPattern)
'    Set xmlElement = xmlElement.cloneNode(True)
'    xmlIndemnityInsurance.replaceChild xmlElement, xmlOmigaElement

    strValue = m_objXmlAssist.GetNodeValue(vxmlAccount, strPattern, False, False)
    
    If Len(Trim$(strValue)) > 0 Then
        Set xmlOmigaElement = xmlMtgeAccountDoc.createElement("INDEMNITYCOMPANYNAME")
        xmlIndemnityInsurance.appendChild xmlOmigaElement
        xmlOmigaElement.Text = strValue
    End If

    strPattern = ".//INDEMNITYMORTGAGEAMOUNT"
'    Set xmlOmigaElement = m_objXmlAssist.GetNode(xmlIndemnityInsurance, strPattern)
'    Set xmlElement = m_objXmlAssist.GetNode(vxmlAccount, strPattern)
'    Set xmlElement = xmlElement.cloneNode(True)
'    xmlIndemnityInsurance.replaceChild xmlElement, xmlOmigaElement
    
    strValue = m_objXmlAssist.GetNodeValue(vxmlAccount, strPattern, False, False)
    
    If Len(Trim$(strValue)) > 0 Then
        Set xmlOmigaElement = xmlMtgeAccountDoc.createElement("INDEMNITYMORTGAGEAMOUNT")
        xmlIndemnityInsurance.appendChild xmlOmigaElement
        xmlOmigaElement.Text = strValue
    End If
    
    strPattern = ".//INDEMNITYAMOUNT"
'    Set xmlOmigaElement = m_objXmlAssist.GetNode(xmlIndemnityInsurance, strPattern)
'    Set xmlElement = m_objXmlAssist.GetNode(vxmlAccount, strPattern)
'    Set xmlElement = xmlElement.cloneNode(True)
'    xmlIndemnityInsurance.replaceChild xmlElement, xmlOmigaElement
    
    strValue = m_objXmlAssist.GetNodeValue(vxmlAccount, strPattern, False, False)
    
    If Len(Trim$(strValue)) > 0 Then
        Set xmlOmigaElement = xmlMtgeAccountDoc.createElement("INDEMNITYAMOUNT")
        xmlIndemnityInsurance.appendChild xmlOmigaElement
        xmlOmigaElement.Text = strValue
    End If
    ' PSC 16/10/2002 BMIDS00575 - End
       
    'BMIDS00006 MDC 19/07/2002 - Customer Account Download
    'Account Relationships ---------------------------------------------------------------------
    Set xmlAccountRelationshipNode = xmlOmigaMtgeAccount.selectSingleNode("ACCOUNTRELATIONSHIPLIST")
    If Not xmlAccountRelationshipNode Is Nothing Then
        xmlOmigaMtgeAccount.removeChild xmlAccountRelationshipNode
    End If
    
    ' PSC 16/10/2002 BMIDS00575 - Start
    Set xmlCustomerListNode = m_objXmlAssist.GetNode(vxmlAccount, "CUSTOMERLIST")

    Set xmlAccountRelationshipNode = xmlMtgeAccountDoc.createElement("ACCOUNTRELATIONSHIPLIST")
    For Each xmlCustomerNode In xmlCustomerListNode.childNodes
        
        ' PSC 16/10/2002 BMIDS00575 - Start
        ' Get the customer version number for this customer
        strCustomerNumber = m_objXmlAssist.GetNodeValue(xmlCustomerNode, "CUSTOMERNUMBER")
        strPattern = ".//CUSTOMERLIST/CUSTOMER[OTHERSYSTEMCUSTOMERNUMBER='" & strCustomerNumber & "']/CUSTOMERVERSIONNUMBER"
        Set xmlCustVersion = m_objXmlAssist.GetNode(vxmlRequest, strPattern)
        
        ' Copy the customer from the Account
        Set xmlElement = xmlCustomerNode.cloneNode(True)
        m_objXmlAssist.ChangeNodeName xmlElement, "CUSTOMER", "ACCOUNTRELATIONSHIP"
        m_objXmlAssist.ChangeNodeName xmlElement, "CUSTOMERNUMBER", "OTHERSYSTEMCUSTOMERNUMBER"
        m_objXmlAssist.ChangeNodeName xmlElement, "OMIGACUSTOMERNUMBER", "CUSTOMERNUMBER"
        
        ' Add in the account guid
        xmlElement.appendChild xmlGUIDNode.cloneNode(True)
        
        ' Add in the customer version
        xmlElement.appendChild xmlCustVersion.cloneNode(True)
        xmlAccountRelationshipNode.appendChild xmlElement
        ' PSC 16/10/2002 BMIDS00575 - End
    Next
    xmlOmigaMtgeAccount.appendChild xmlAccountRelationshipNode
    'BMIDS00006 MDC 19/07/2002 - End
    
    ' Update the mortgage account
    xmlRequestElem.appendChild xmlOmigaMtgeAccount

    strResponse = objCustomerFinancialBO.UpdateMortgageAccount(xmlMtgeAccountDoc.xml)
    Set xmlResponseDoc = m_objXmlAssist.load(strResponse, TypeName(Me), cstrFunctionName)
    m_objErrAssist.CheckResponse xmlResponseDoc.documentElement, vxmlResponse, True
    
    Set UpdateMortgageAccount = m_objXmlAssist.GetNode(xmlMtgeAccountDoc, ".//ACCOUNTGUID")
    
    Set xmlMtgeAccountDoc = Nothing
    Set xmlResponseDoc = Nothing
    Set xmlRequestElem = Nothing
    Set xmlElement = Nothing
    Set xmlOmigaElement = Nothing
    Set xmlOmigaMtgeAccount = Nothing
    Set xmlTableElem = Nothing
    Set objCustomerFinancialBO = Nothing
    Set xmlAccount = Nothing
    Set xmlIndemnityInsurance = Nothing

    'BMIDS00006 MDC 22/07/2002 - Customer Account Download
    Set xmlGUIDNode = Nothing
    'BMIDS00682 MDC 24/10/2002 - No longer required
    'Set xmlTempRequestElem = Nothing
    'Set objCustomerBO = Nothing
    'BMIDS00682 MDC 24/10/2002 - End
    Set xmlAccountRelationshipNode = Nothing
    Set xmlCustomerListNode = Nothing
    Set xmlCustomerNode = Nothing
    'BMIDS00006 MDC 22/07/2002 - End
    
    Set xmlCustVersion = Nothing         ' PSC 16/10/2002 BMIDS00575


    
    Exit Function
                                                
UpdateMortgageAccountVbErr:

    Set xmlMtgeAccountDoc = Nothing
    Set xmlResponseDoc = Nothing
    Set xmlRequestElem = Nothing
    Set xmlElement = Nothing
    Set xmlOmigaElement = Nothing
    Set xmlOmigaMtgeAccount = Nothing
    Set xmlTableElem = Nothing
    Set objCustomerFinancialBO = Nothing
    Set xmlAccount = Nothing
    Set xmlIndemnityInsurance = Nothing
    
    
    m_objErrAssist.AddToErrSource cstrFunctionName
    
    Err.Raise Err.Number, Err.Source, Err.Description
    
End Function


Private Sub UpdateMortgageLoan(ByVal vxmlRequest As IXMLDOMElement, _
                               ByVal vxmlLoan As IXMLDOMElement, _
                               ByVal vxmlOmigaLoan As IXMLDOMElement, _
                               ByVal vxmlAccountGuid As IXMLDOMElement, _
                               ByVal vxmlResponse As IXMLDOMElement)

' header ----------------------------------------------------------------------------------
' description:  Creates the mortgage loan from the details off the administration system
' pass:         vxmlRequest                     xml containing Request
'               vxmlLoan                        xml containing Loan details
'               xmlAccountGuid                  xml holding Guid ogf the account the loan is
'                                               attached to
'               vxmlResponse                    xml to attach warnings to
' return:
'------------------------------------------------------------------------------------------

On Error GoTo UpdateMortgageLoanVbErr

    Const cstrFunctionName As String = "UpdateMortgageLoan"
    
    Dim xmlMtgeLoanDoc As FreeThreadedDOMDocument40
    Dim xmlResponseDoc As FreeThreadedDOMDocument40
    Dim xmlRequestElem As IXMLDOMElement
    Dim xmlElement As IXMLDOMElement
    Dim xmlOmigaElement As IXMLDOMElement
    Dim xmlOmigaLoan As IXMLDOMElement
    Dim objCustomerFinancialBO As CustomerFinancialBO
    
    Dim strResponse As String
    Dim strPattern As String
    
    Set objCustomerFinancialBO = _
                    m_objContext.CreateInstance(gstrCUSTOMER_FINANCIAL & ".CustomerFinancialBO")
   
    Set xmlOmigaLoan = vxmlOmigaLoan.cloneNode(True)
    
    Set xmlMtgeLoanDoc = New FreeThreadedDOMDocument40
    xmlMtgeLoanDoc.validateOnParse = False
    xmlMtgeLoanDoc.setProperty "NewParser", True
    xmlMtgeLoanDoc.async = False
    
    Set xmlRequestElem = m_objXmlAssist.GetRequestNodeEx(vxmlRequest)
    xmlMtgeLoanDoc.appendChild xmlRequestElem
    
    
    'PSC 16/10/2002 BMIDS00575 - Start
    strPattern = ".//LOANACCOUNTNUMBER"
    Set xmlElement = m_objXmlAssist.GetNode(vxmlLoan, strPattern)
    
    If Len(Trim$(xmlElement.Text)) > 0 Then
        Set xmlOmigaElement = m_objXmlAssist.GetNode(xmlOmigaLoan, strPattern)
        Set xmlElement = xmlElement.cloneNode(True)
        xmlOmigaLoan.replaceChild xmlElement, xmlOmigaElement
    End If

    strPattern = ".//MORTGAGEPRODUCTCODE"
    Set xmlElement = m_objXmlAssist.GetNode(vxmlLoan, strPattern)

    If Len(Trim$(xmlElement.Text)) > 0 Then
        Set xmlOmigaElement = m_objXmlAssist.GetNode(xmlOmigaLoan, strPattern)
        Set xmlElement = xmlElement.cloneNode(True)
        xmlOmigaLoan.replaceChild xmlElement, xmlOmigaElement
    End If
   
    strPattern = ".//MORTGAGEPRODUCTDESCRIPTION"
    Set xmlElement = m_objXmlAssist.GetNode(vxmlLoan, strPattern)

    If Len(Trim$(xmlElement.Text)) > 0 Then
        Set xmlOmigaElement = m_objXmlAssist.GetNode(xmlOmigaLoan, strPattern)
        Set xmlElement = xmlElement.cloneNode(True)
        xmlOmigaLoan.replaceChild xmlElement, xmlOmigaElement
    End If
    
    strPattern = ".//INTERESTRATE"
    Set xmlElement = m_objXmlAssist.GetNode(vxmlLoan, strPattern)
    
    If Len(Trim$(xmlElement.Text)) > 0 Then
        Set xmlOmigaElement = m_objXmlAssist.GetNode(xmlOmigaLoan, strPattern)
        Set xmlElement = xmlElement.cloneNode(True)
        xmlOmigaLoan.replaceChild xmlElement, xmlOmigaElement
    End If
    
    strPattern = ".//PURPOSEOFLOAN"
    Set xmlElement = m_objXmlAssist.GetNode(vxmlLoan, strPattern)
    
    If Len(Trim$(xmlElement.Text)) > 0 Then
        Set xmlOmigaElement = m_objXmlAssist.GetNode(xmlOmigaLoan, strPattern)
        Set xmlElement = xmlElement.cloneNode(True)
        xmlOmigaLoan.replaceChild xmlElement, xmlOmigaElement
    End If
       
    strPattern = ".//OUTSTANDINGBALANCE"
    Set xmlElement = m_objXmlAssist.GetNode(vxmlLoan, strPattern)
    
    If Len(Trim$(xmlElement.Text)) > 0 Then
        Set xmlOmigaElement = m_objXmlAssist.GetNode(xmlOmigaLoan, strPattern)
        Set xmlElement = xmlElement.cloneNode(True)
        xmlOmigaLoan.replaceChild xmlElement, xmlOmigaElement
    End If
    
    strPattern = ".//REPAYMENTTYPE"
    Set xmlElement = m_objXmlAssist.GetNode(vxmlLoan, strPattern)
    
    If Len(Trim$(xmlElement.Text)) > 0 Then
        Set xmlOmigaElement = m_objXmlAssist.GetNode(xmlOmigaLoan, strPattern)
        Set xmlElement = xmlElement.cloneNode(True)
        xmlOmigaLoan.replaceChild xmlElement, xmlOmigaElement
    End If
    
    strPattern = ".//ORIGINALLOANAMOUNT"
    Set xmlElement = m_objXmlAssist.GetNode(vxmlLoan, strPattern)
    
    If Len(Trim$(xmlElement.Text)) > 0 Then
        Set xmlOmigaElement = m_objXmlAssist.GetNode(xmlOmigaLoan, strPattern)
        Set xmlElement = xmlElement.cloneNode(True)
        xmlOmigaLoan.replaceChild xmlElement, xmlOmigaElement
    End If
    
    'BMIDS756 GHun
    strPattern = ".//ORIGINALPARTANDPARTINTONLYAMT"
    Set xmlElement = m_objXmlAssist.GetNode(vxmlLoan, strPattern)
    
    If Len(Trim$(xmlElement.Text)) > 0 Then
        Set xmlOmigaElement = m_objXmlAssist.GetNode(xmlOmigaLoan, strPattern)
        Set xmlElement = xmlElement.cloneNode(True)
        xmlOmigaLoan.replaceChild xmlElement, xmlOmigaElement
    End If
    'BMIDS756 End
    
    strPattern = ".//MONTHLYREPAYMENT"
    Set xmlElement = m_objXmlAssist.GetNode(vxmlLoan, strPattern)
    
    If Len(Trim$(xmlElement.Text)) > 0 Then
        Set xmlOmigaElement = m_objXmlAssist.GetNode(xmlOmigaLoan, strPattern)
        Set xmlElement = xmlElement.cloneNode(True)
        xmlOmigaLoan.replaceChild xmlElement, xmlOmigaElement
    End If
    
    strPattern = ".//ORIGINALTERMYEARS"
    Set xmlElement = m_objXmlAssist.GetNode(vxmlLoan, strPattern)
    
    If Len(Trim$(xmlElement.Text)) > 0 Then
        Set xmlOmigaElement = m_objXmlAssist.GetNode(xmlOmigaLoan, strPattern)
        Set xmlElement = xmlElement.cloneNode(True)
        xmlOmigaLoan.replaceChild xmlElement, xmlOmigaElement
    End If
    
    strPattern = ".//ORIGINALTERMMONTHS"
    Set xmlElement = m_objXmlAssist.GetNode(vxmlLoan, strPattern)
    
    If Len(Trim$(xmlElement.Text)) > 0 Then
        Set xmlOmigaElement = m_objXmlAssist.GetNode(xmlOmigaLoan, strPattern)
        Set xmlElement = xmlElement.cloneNode(True)
        xmlOmigaLoan.replaceChild xmlElement, xmlOmigaElement
    End If
    
    strPattern = ".//STARTDATE"
    Set xmlElement = m_objXmlAssist.GetNode(vxmlLoan, strPattern)
    
    If Len(Trim$(xmlElement.Text)) > 0 Then
        Set xmlOmigaElement = m_objXmlAssist.GetNode(xmlOmigaLoan, strPattern)
        Set xmlElement = xmlElement.cloneNode(True)
        xmlOmigaLoan.replaceChild xmlElement, xmlOmigaElement
    End If
    
    ' PSC 16/10/2002 BMIDS00575
    ' Comment out code
    '*=[MC]BMIDS756 Regulation Changes Start
    strPattern = ".//REDEMPTIONSTATUS"
    Set xmlOmigaElement = m_objXmlAssist.GetNode(xmlOmigaLoan, strPattern)
    Set xmlElement = m_objXmlAssist.GetNode(vxmlLoan, strPattern)
    Set xmlElement = xmlElement.cloneNode(True)
    xmlOmigaLoan.replaceChild xmlElement, xmlOmigaElement

    strPattern = ".//REDEMPTIONDATE"
    Set xmlOmigaElement = m_objXmlAssist.GetNode(xmlOmigaLoan, strPattern)
    Set xmlElement = m_objXmlAssist.GetNode(vxmlLoan, strPattern)
    Set xmlElement = xmlElement.cloneNode(True)
    xmlOmigaLoan.replaceChild xmlElement, xmlOmigaElement
    '*=[MC]BMIDS756 Regulation Changes End
    ' PSC 16/10/2002 BMIDS00575 - End
    
    'BMIDS907 GHun
    strPattern = ".//PRODUCTSTEP"
    Set xmlElement = m_objXmlAssist.GetNode(vxmlLoan, strPattern)
    Set xmlOmigaElement = m_objXmlAssist.GetNode(xmlOmigaLoan, strPattern)
    xmlOmigaLoan.replaceChild xmlElement, xmlOmigaElement
        
    strPattern = ".//REMAININGSTEPDURATION"
    Set xmlElement = m_objXmlAssist.GetNode(vxmlLoan, strPattern)
    Set xmlOmigaElement = m_objXmlAssist.GetNode(xmlOmigaLoan, strPattern)
    xmlOmigaLoan.replaceChild xmlElement, xmlOmigaElement
    
    strPattern = ".//REMAININGCAPITALINTERESTAMOUNT"
    Set xmlElement = m_objXmlAssist.GetNode(vxmlLoan, strPattern)
    Set xmlOmigaElement = m_objXmlAssist.GetNode(xmlOmigaLoan, strPattern)
    xmlOmigaLoan.replaceChild xmlElement, xmlOmigaElement
    
    strPattern = ".//REMAININGINTERESTONLYAMOUNT"
    Set xmlElement = m_objXmlAssist.GetNode(vxmlLoan, strPattern)
    Set xmlOmigaElement = m_objXmlAssist.GetNode(xmlOmigaLoan, strPattern)
    xmlOmigaLoan.replaceChild xmlElement, xmlOmigaElement
    'BMIDS907 End
    
    xmlRequestElem.appendChild xmlOmigaLoan
    
    ' Create the mortgage loan
    strResponse = objCustomerFinancialBO.UpdateMortgageLoan(xmlMtgeLoanDoc.xml)
    Set xmlResponseDoc = m_objXmlAssist.load(strResponse, TypeName(Me), cstrFunctionName)
    m_objErrAssist.CheckResponse xmlResponseDoc.documentElement, vxmlResponse, True
    
    Set xmlMtgeLoanDoc = Nothing
    Set xmlResponseDoc = Nothing
    Set xmlRequestElem = Nothing
    Set xmlElement = Nothing
    Set xmlOmigaElement = Nothing
    Set xmlOmigaLoan = Nothing
    Set objCustomerFinancialBO = Nothing
    
    Exit Sub
    
UpdateMortgageLoanVbErr:

    Set xmlMtgeLoanDoc = Nothing
    Set xmlResponseDoc = Nothing
    Set xmlRequestElem = Nothing
    Set xmlElement = Nothing
    Set xmlOmigaElement = Nothing
    Set xmlOmigaLoan = Nothing
    Set objCustomerFinancialBO = Nothing
    
    m_objErrAssist.AddToErrSource cstrFunctionName
    
    Err.Raise Err.Number, Err.Source, Err.Description
    
End Sub


Private Sub CreateMortgageLoan(ByVal vxmlRequest As IXMLDOMElement, _
                               ByVal vxmlLoan As IXMLDOMElement, _
                               ByVal vxmlAccountGuid As IXMLDOMElement, _
                               ByVal vxmlResponse As IXMLDOMElement)

' header ----------------------------------------------------------------------------------
' description:  Creates the mortgage loan from the details off the administration system
' pass:         vxmlRequest                     xml containing Request
'               vxmlLoan                        xml containing Loan details
'               vxmlResponse                    xml to attach warnings to
' return:       ImportAccountAndLoans           Guid of the created account
'------------------------------------------------------------------------------------------

On Error GoTo CreateMortgageLoanLoanVbErr

    Const cstrFunctionName As String = "CreateMortgageLoan"

    
    Dim xmlMtgeLoanDoc As FreeThreadedDOMDocument40
    Dim xmlRequestElem As IXMLDOMElement
    Dim xmlElement As IXMLDOMElement
    Dim xmlTableElem As IXMLDOMElement
    Dim xmlResponseDoc As FreeThreadedDOMDocument40
    Dim objCustomerFinancialBO As CustomerFinancialBO
    
    Dim strResponse As String
    
    Set objCustomerFinancialBO = _
                    m_objContext.CreateInstance(gstrCUSTOMER_FINANCIAL & ".CustomerFinancialBO")
       
    Set xmlMtgeLoanDoc = New FreeThreadedDOMDocument40
    xmlMtgeLoanDoc.validateOnParse = False
    xmlMtgeLoanDoc.setProperty "NewParser", True
    xmlMtgeLoanDoc.async = False
    
    Set xmlRequestElem = m_objXmlAssist.GetRequestNodeEx(vxmlRequest)
    xmlMtgeLoanDoc.appendChild xmlRequestElem
    
    Set xmlTableElem = vxmlLoan.cloneNode(True)
    xmlRequestElem.appendChild xmlTableElem
    xmlTableElem.appendChild vxmlAccountGuid.cloneNode(True)
        
    ' Create the mortgage loan
    strResponse = objCustomerFinancialBO.CreateMortgageLoan(xmlMtgeLoanDoc.xml)
    Set xmlResponseDoc = m_objXmlAssist.load(strResponse, TypeName(Me), cstrFunctionName)
    m_objErrAssist.CheckResponse xmlResponseDoc.documentElement, vxmlResponse, True
    
    Set xmlMtgeLoanDoc = Nothing
    Set xmlResponseDoc = Nothing
    Set xmlRequestElem = Nothing
    Set xmlElement = Nothing
    Set xmlTableElem = Nothing
    Set objCustomerFinancialBO = Nothing
    
    
    Exit Sub
                                                
CreateMortgageLoanLoanVbErr:

    Set xmlMtgeLoanDoc = Nothing
    Set xmlResponseDoc = Nothing
    Set xmlRequestElem = Nothing
    Set xmlElement = Nothing
    Set xmlTableElem = Nothing
    Set objCustomerFinancialBO = Nothing
    
    
    m_objErrAssist.AddToErrSource cstrFunctionName
    
    Err.Raise Err.Number, Err.Source, Err.Description
    
End Sub

Private Sub ImportArrearsHistory(ByVal vxmlRequest As IXMLDOMElement, _
                                 ByVal vxmlAccount As IXMLDOMElement, _
                                 ByVal vstrAccountGuid As String, _
                                 ByVal vxmlResponse As IXMLDOMElement)
' header ----------------------------------------------------------------------------------
' description:  Deletes the existing arrears on the account and inserts the new arrears
' pass:         vxmlRequest                     xml containing Request
'               vxmlAccount                     xml containing Account details
'               vstrAccountGuid                 Guid of the account in Omiga
'               vxmlResponse                    xml to attach warnings to
' return:
' History
' RF    27/07/2004  BMIDS818    Rewrite to call new stored procedure usp_DeleteArrearsHistoryForAccount.
'                               Removed old commented code.
'------------------------------------------------------------------------------------------
On Error GoTo ImportArrearsHistoryVbErr

    Const cstrFunctionName As String = "ImportArrearsHistory"
    
    '------------------------------------------------------------------------------------------
    ' Delete the existing arrears for this account
    '------------------------------------------------------------------------------------------
    
    DeleteExistingArrears vstrAccountGuid
    
    '------------------------------------------------------------------------------------------
    ' Create any new arrears for this account
    '------------------------------------------------------------------------------------------
    
    Dim strResponse As String
    Dim strData As String
    Dim strDescOfLoan As String
    Dim elemArrearsHist As IXMLDOMElement
    Dim elemRequest As IXMLDOMElement
    Dim elemTemp As IXMLDOMElement
    Dim objCustomerFinancialBO As CustomerFinancialBO
    
    'SYS4703 - ARREARSHISTORY is not mandatory.
    Set elemArrearsHist = m_objXmlAssist.GetNode(vxmlAccount, ".//ARREARSHISTORY", False)
   
    ' If we have arrears on this account from the admin system then create them in Omiga
    If Not elemArrearsHist Is Nothing Then
        
        'BMIDS00447 Check if there is any arrears data to save
        
        strDescOfLoan = Trim$(xmlGetNodeText(elemArrearsHist, "DESCRIPTIONOFLOAN"))
        
        strData = xmlGetNodeText(elemArrearsHist, "MAXIMUMBALANCE") & _
            xmlGetNodeText(elemArrearsHist, "MAXIMUMNUMBEROFMONTHS") & _
            xmlGetNodeText(elemArrearsHist, "DATECLEARED") & _
            strDescOfLoan & _
            xmlGetNodeText(elemArrearsHist, "CURRENTYEARSINARREARS") & _
            xmlGetNodeText(elemArrearsHist, "LASTTWOYEARSINARREARS")
        
        If Len(Trim(strData)) > 0 Then
            
            'Xml contains some data. Further validation will be performed by CustomerFinancialBO
            
            'BMIDS00447 Default description of loan if it is blank
            If Len(strDescOfLoan) = 0 Then
                elemArrearsHist.selectSingleNode("DESCRIPTIONOFLOAN").Text = _
                    GetFirstComboValueId("ArrearsLoanType", "M")
            End If
            'BMIDS00447 End
            
            Set elemTemp = elemArrearsHist.ownerDocument.createElement("ACCOUNTGUID")
            elemTemp.Text = vstrAccountGuid
            elemArrearsHist.appendChild elemTemp
            
            Set elemRequest = m_objXmlAssist.GetRequestNodeEx(vxmlRequest)
            elemRequest.appendChild elemArrearsHist.cloneNode(True)
            
            Set objCustomerFinancialBO = _
                m_objContext.CreateInstance(gstrCUSTOMER_FINANCIAL & ".CustomerFinancialBO")

            strResponse = objCustomerFinancialBO.CreateArrearsHistory(elemRequest.xml)
            m_objErrAssist.CheckXMLResponse strResponse, True, vxmlResponse
        End If
    End If
      
    Set elemArrearsHist = Nothing
    Set elemRequest = Nothing
    Set elemTemp = Nothing
    Set objCustomerFinancialBO = Nothing
    
    Exit Sub
                                                
ImportArrearsHistoryVbErr:
    
    Set elemArrearsHist = Nothing
    Set elemRequest = Nothing
    Set elemTemp = Nothing
    Set objCustomerFinancialBO = Nothing
    
    m_objErrAssist.AddToErrSource cstrFunctionName
    
    Err.Raise Err.Number, Err.Source, Err.Description

End Sub

Private Sub DeleteExistingArrears(ByVal vstrAccountGuid As String)
' header ----------------------------------------------------------------------------------
' description:  Deletes the existing arrears on the account
' pass:         vstrAccountGuid     Guid of the account in Omiga
' return:
' History
' RF    27/07/2004  BMIDS818    Created
'------------------------------------------------------------------------------------------
On Error GoTo DeleteExistingArrearsVbErr

    Const cstrFunctionName As String = "DeleteExistingArrears"
    
    Dim cmd As ADODB.Command
    Dim objADOAssist As IADOAssist
    Dim bytAcctGuid() As Byte
    Dim objSqlAssist As SQLAssist
    
    Set cmd = New ADODB.Command
    Set objADOAssist = m_objContext.CreateInstance(gstrBASE_COMPONENT & ".ADOAssist")
    
    With cmd
        .ActiveConnection = objADOAssist.GetConnStr
        Set objSqlAssist = New SQLAssist
        bytAcctGuid = objSqlAssist.GuidStringToByteArray(vstrAccountGuid)
        .CommandText = "usp_DeleteArrearsHistoryForAccount"
        .Parameters.Append .CreateParameter( _
            "@p_AccountGuid", _
            adVarBinary, _
            adParamInput, _
            UBound(bytAcctGuid) + 1, _
            bytAcctGuid)
        .Execute , , adCmdStoredProc
    End With
    
    Set objSqlAssist = Nothing
    Set cmd = Nothing
    Set objADOAssist = Nothing
    
    Exit Sub
                                                
DeleteExistingArrearsVbErr:

    Set objSqlAssist = Nothing
    Set cmd = Nothing
    Set objADOAssist = Nothing
    
    m_objErrAssist.AddToErrSource cstrFunctionName
    
    Err.Raise Err.Number, Err.Source, Err.Description
    
End Sub

'Private Function GetOmigaCustomer(ByVal xmlCustomerNode As IXMLDOMNode) As IXMLDOMNode
'' header ----------------------------------------------------------------------------------
'' procedure:    GetOmigaCustomer
'' author:       Mark Coleman
'' date:         16/07/2002
'' description:  Finds an Omiga Customer record from
''               a) OmigaCustomerNumber
''               b) OtherSystemCustomerNumber
'' pass:         xmlCustomerNode  XML Input request specifying the customer to find
'' return:       Customer XML for the customer found. Nothing if not found
''------------------------------------------------------------------------------------------
'On Error GoTo GetOmigaCustomer_Exit
'
'Dim strFunctionName As String
'strFunctionName = "GetOmigaCustomer"
'
'Dim xmlDoc As FreeThreadedDOMDocument40
'Dim xmlRequest As IXMLDOMNode
'Dim xmlTempCustomerNode As IXMLDOMNode
'Dim xmlCustomerNumberNode As IXMLDOMNode
'Dim xmlReturnCustomer As IXMLDOMNode
'
'Dim objCustomerBO As CustomerBO
'
'Dim strOmigaCustomerNumber As String
'Dim strAdminCustomerNumber As String
'Dim strResponse As String
'Dim lngErr As Long
'
'    ' <VSA> Visual Studio Analyser Support
'    #If USING_VSA Then
'        Dim VSA As New vsa_shared: VSA.Initialize (TypeName(Me) & "." & strFunctionName)
'    #End If
'
'    strOmigaCustomerNumber = xmlGetNodeText(xmlCustomerNode, ".//OMIGACUSTOMERNUMBER")
'    strAdminCustomerNumber = xmlGetNodeText(xmlCustomerNode, ".//CUSTOMERNUMBER")
'
'    If Len(Trim$(strAdminCustomerNumber)) = 0 And Len(Trim$(strOmigaCustomerNumber)) = 0 Then
'        Set GetOmigaCustomer = Nothing
'    Else
'        Set objCustomerBO = m_objContext.CreateInstance(gstrCUSTOMER_COMPONENT & ".CustomerBO")
'
'        Set xmlDoc = New FreeThreadedDOMDocument40
'        xmlDoc.async = False
'
'        Set xmlRequest = xmlDoc.createElement("REQUEST")
'        Set xmlTempCustomerNode = xmlDoc.createElement("CUSTOMER")
'        xmlRequest.appendChild xmlTempCustomerNode
'        Set xmlCustomerNumberNode = xmlDoc.createElement("CUSTOMERNUMBER")
'        xmlTempCustomerNode.appendChild xmlCustomerNumberNode
'
'        'If Omiga Customer Number exists, try that first
'        If Len(Trim$(strOmigaCustomerNumber)) > 0 Then
'            xmlCustomerNumberNode.Text = strOmigaCustomerNumber
'            strResponse = objCustomerBO.GetCustomerDetails(xmlRequest.xml)
'            lngErr = errCheckXMLResponse(strResponse, False)
'            If lngErr <> 0 And lngErr <> oeRecordNotFound Then
'                errThrowError strFunctionName, lngErr
'            ElseIf lngErr = 0 Then
'                'Found Customer
'                Set xmlDoc = xmlLoad(strResponse, strFunctionName)
'                Set xmlReturnCustomer = xmlDoc.selectSingleNode(".//CUSTOMER")
'            End If
'        End If
'
'        'Otherwise try the Admin System Number
'        If xmlReturnCustomer Is Nothing And Len(Trim$(strOmigaCustomerNumber)) > 0 Then
'            xmlCustomerNumberNode.Text = strAdminCustomerNumber
'            strResponse = objCustomerBO.GetCustomerFromLegacyNumber(xmlRequest.xml)
'            lngErr = errCheckXMLResponse(strResponse, False)
'            If lngErr <> 0 And lngErr <> oeRecordNotFound Then
'                errThrowError strFunctionName, lngErr
'            ElseIf lngErr = 0 Then
'                'Found Customer
'                Set xmlDoc = xmlLoad(strResponse, strFunctionName)
'                Set xmlReturnCustomer = xmlDoc.selectSingleNode(".//CUSTOMER")
'            End If
'        End If
'    End If
'
'    Set GetOmigaCustomer = xmlReturnCustomer
'
'GetOmigaCustomer_Exit:
'    Set xmlDoc = Nothing
'    Set xmlRequest = Nothing
'    Set xmlTempCustomerNode = Nothing
'    Set xmlCustomerNumberNode = Nothing
'    Set xmlReturnCustomer = Nothing
'    Set objCustomerBO = Nothing
'
'    ' <VSA> Visual Studio Analyser Support
'    #If USING_VSA Then
'        Set VSA = Nothing
'    #End If
'
'    errCheckError strFunctionName, TypeName(Me)
'
'End Function

'BMIDS00566
Private Sub UpdateAppCorresSalutation(ByVal strApplicationNumber As String, _
                                                      ByVal intAFFNumber As Integer)
'--------------------------------------------------------------------------------------------
'Description : Maintains the Correspondence salutation whenever the customers for an
'              application change
'--------------------------------------------------------------------------------------------
    Const cstrFunctionName As String = "UpdateAppCorresSalutation"
    Dim objIADOAssist   As IADOAssist
    Dim cmd             As ADODB.Command
    Dim Cn              As ADODB.Connection
    Dim lngErrNo        As Long
    
On Error GoTo UpdateAppCorresSalutationVbErr
    
    ' Create the Command object, assign values to the parameters and execute it
    Set Cn = New ADODB.Connection
    If m_objContext Is Nothing Then
        Set objIADOAssist = New ADOAssist
    Else
        Set objIADOAssist = m_objContext.CreateInstance(gstrBASE_COMPONENT & ".ADOAssist")
    End If
    Cn.ConnectionString = objIADOAssist.GetConnStr
    Cn.CursorLocation = adUseClient
    Cn.open
    
    Set cmd = New ADODB.Command
    With cmd
        Set .ActiveConnection = Cn
        .CommandType = adCmdStoredProc
        'ASt 15/03/01 CC012 Compatible with SQL Server and Oracle
    #If GENERIC_SQL Then
        If objIADOAssist.GetDBEngine = Oracle Then
            .CommandText = "sp_CustomerReg.UpdateCorrespondenceSalutation"
        ElseIf objIADOAssist.GetDBEngine = SQLServer Then
            .CommandText = "usp_CustomerRegUpdateCorrespondenceSalutation"
        End If
    #Else
        .CommandText = "sp_CustomerReg.UpdateCorrespondenceSalutation"
    #End If
        .Parameters.Append .CreateParameter(, adBSTR, adParamInput, 12, strApplicationNumber)
        .Parameters.Append .CreateParameter(, adInteger, adParamInput, 5, intAFFNumber)
        .Execute
    End With
    '   Check for any errors raised
    
    If cmd.ActiveConnection.Errors.Count <> 0 Then
        lngErrNo = cmd.ActiveConnection.Errors.Item(0).Number
        m_objErrAssist.ThrowError TypeName(Me), cstrFunctionName, Err.Number
    End If
    
    Set cmd.ActiveConnection = Nothing
    Cn.Close
    Set Cn = Nothing
    Set cmd = Nothing
    If Not m_objContext Is Nothing Then
        m_objContext.SetComplete
    End If
    Exit Sub
    
UpdateAppCorresSalutationVbErr:
    ' Close the connection if still open
    If Not Cn Is Nothing Then
        If Cn.State = adStateOpen Then
            Cn.Close
        End If
    End If
    
    Set Cn = Nothing
    Set cmd = Nothing
    
    Err.Raise Err.Number, Err.Source, Err.Description
End Sub
'BMIDS00566 End

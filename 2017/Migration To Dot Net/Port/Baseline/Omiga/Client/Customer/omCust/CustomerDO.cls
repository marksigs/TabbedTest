VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 3  'UsesTransaction
END
Attribute VB_Name = "CustomerDO"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True

'Workfile:      CustomerDO.cls
'Copyright:     Copyright © 1999 Marlborough Stirling

'Description:   Code for person Data Object
'               contains standard interface:
'               Public Sub Create(vstrXMLRequest As String)
'               Public Sub Update(vstrXMLRequest As String)
'               Public Sub Delete(vstrXMLRequest As String)
'               Public Function GetData(vstrXMLRequest As String) As String
'Dependencies:  CustomerBO, CustomerTxBO
'
'Issues:        Instancing:         MultiUse
'               MTSTransactionMode: UsesTransaction
'------------------------------------------------------------------------------------------
'History:
'
'Prog   Date     Description
'SR    04/11/99  Created
'SR    21/02/00  Modified method CreateCustomerDetails - Assign value to channelId
'SR    28/02/00  Added validation of telephone in WriteCustomerTelephones
'SR    03/03/00  Modified method DeleteCustomerAddress AQR - SYS0392
'SR    08/03/00  Added sorting order to 'FindCustomerAddressList' AQR - SYS0390
'SR    10/03/00  CreateLock uses the search pattren './/' rather than '//' to search
'                for CustomerNumber in the node
'MH    14/03/00  GenerateNextCustomerNumber no longer uses Stored Procedure &
'                only returns a checkdigit friendly number. This should save
'                significant time on creating connections and queries.
'SR    15/03/00  AQR/SYS0220 - Modified method 'UpdateCustomerDetails'. Check whether the
'                'Customer' table is to be updated or not.
'SR    02/05/00  AQR/SYS0227 - Modified method FindCustomer - Formatting the values assigned
'                to Parameters of the Command Object used to call stored procedure.
'SR    02/05/00  New method UpdateCustomerVersion
'MC    02/06/00  Added Tax Details option to UpdateCustomerVersion
'SR     02/06/00 Added an optional parameter to method 'FindCustomerAddressList' used to sort
'                   the output in the order of CustomerAddressSequenceNumber
'JLD    06/07/00 in FindDependantList, don't pass in an incorrect search tag for FindListMultiple
'                as it results in a 'select * from' statement.
'BG     26/07/00 Added CreatePropertyInsuranceDetails which calls CreateCurrentProperty and
'                CreateHomeInsurance.
'BG     27/07/00 Added GetPropertyInsuranceDetails method
'SR     01/09/00 Modified DeleteCustomerAddress - Added an additional input param
'                'DeleteMortgageAccount'; if this is true, delete the linked records from MortageAccount
'IVW    27/09/00 Changed SQLAssist objects to not be created under MTS.
'LD     07/11/00 Explicity close database connections
'LD     07/11/00 Explicity close recordsets
'LD     07/11/00 Explicity destroy command objects
'CL     20/12/00 SYS1695 - Amend GetNextCustomerNumber to disconnect the recordset before closing the
'                          connection
'CL     21/12/00 SYS1755 Changes made to implement Contact History
'DJP    03/01/00 SYS1779 Change FindCustomer to disconnect recordset before closing connection.
'PSC    09/01/01 SYS1750 Amend FindCustomer to take new inputs and return extra output
'                        Add GetCustomerFromLegacyNumber()
'DM     14/02/01 SYS1946 Changed ICustomerDO_GetNextCustomerNumber to use stored procs instead of sequences
'                        Note this is conditionally compiled out.
'AS     30/05/01 CC012  Ported to SQL Server
'SR     05/06/01 SYS2325 New method 'UpdateCustomer'
'DM     06/06/01 SYS1946 Made change to stored proc name in GetNextCustomerNumber
'SR     12/06/01 SYS2362 Modified method 'GetXMLFromfindList'. Added 'SOURCE' node to the response
'LD     11/06/01 SYS2367 SQL Server Port - Length must be specified in calls to CreateParameter
'LD     18/06/01 SYS2368 - Now using USP_GETNEXTCUSTOMERNUMBER stored procedure
'DRC    26/06/01 SYS     modified method AddSalutationToRequest and ensure that update customer (personal)
'                        details uses this. Also changed referennces to GUIDAssist class to function
'JR     28/06/01 SYSOmiplus24 - Modified method IsEmptyTelephoneNode to include CountryCode and AreaCode
'AS     04/07/01 CC012  SQL Server usp_CustomerRegNoRowsReturned stored procedure no longer called as
'                       usp_CustomerRegFindCustomerDetails returns the total number of matching rows.
'DM     20/07/01 SYS2511 Changed parameter type to adVarChar
'JR     24/07/01 SYS2505 Amend IsEmptyTelephoneNode, comment out Country/Area Code
'JR     24/08/01 Omiplus24 Modified method IsEmptyTelephoneNode to re-add CountryCode and AreaCode
'AD     19/10/01 Removed the MoveFirst as it's a ForwardOnly cursor. SYS2718.
'JR     26/10/01 SYS2860 Comment out: References to CountryCode/Areacode in IsEmptyTelephoneNode until further notice
'JR     03/12/01 SYS2459 re-apply Omiplus24 telephone changes
'SG     30/05/02 SYS4767 MSMS to Core integration
'----------------------------------------------------------------------------------------------------------
'MARS Specific History:
'
'Prog   Date        Description
'PSC    28/06/2005  MAR5    Add GetCustomerAddressListAndRole function.
'HM     08/09/2005  MAR29   GetCustomerDetails has been extended to include wrap up details
'HM     08/09/2005  MAR29   ICustomerTxBO_Create/Update/DeleteWrapUpDetails were added
'PSC    10/10/2005  MAR57   Added FindCustomerCRSData
'MV     20/10/2005  MAR96   Amended ICustomerDO_CreateCustomerTelephone()
'HMA    12/10/2005  MAR23   Add GetCustomerNumbers function. Added CreateKYCCheck.
'PSC    16/01/2006  MAR1050 Amend DeleteAreasOfInterest, DeleteCustomerVerifications and
'                           DeleteCustomerTelephones to clear any errors returned.
'PSC    03/04/2006  MAR1573 Remove code for CreateKYCCheck and use CRUD instead. Method
'                           kept for Binary Compatibility
'------------------------------------------------------------------------------------------
'EPSOM Specific History:
'
'Prog   Date        Description
'SAB    21/04/2006  EP399   Updates CustomerAddress.lastamendeddated when a change is made to
'                           the relevant address
'IK     17/05/2006  EP561   add IsCheckDigitRequired test
'PE     08/06/2006  EP9734  Mars merge/MAR1848 Changed GetCustomerFromLegacyNumber to improve performance
'------------------------------------------------------------------------------------------
Option Explicit

Implements ICustomerDO
Implements ObjectControl

Private m_objContext As ObjectContext
Private m_objErrAssist As ErrAssist
Private m_objXmlAssist As XMLAssist             ' Remove if not required
Private m_objIDOAssist As IDOAssist

Const cstrPERSON_COMPONENT = "omCust"
Const cstrADDRESS_COMPONENT = "omCust"
Const cstrTHIRDPARTY_COMPONENT = "omTp"
Private Const cintFINDCUSTOMERPAGE = 10

Private Function ICustomerDO_CreateContactHistory(ByVal vxmlTableElement As IXMLDOMElement) As IXMLDOMNode
' header ----------------------------------------------------------------------------------
' description:
'   create an instance of the persistant data associated with this data object
'   for each set of data in the request
' pass:
'   vxmlTableElement  xml element containing the request
' return:                n/a
'------------------------------------------------------------------------------------------
On Error GoTo CreateContactHistoryVbErr

    Const strFunctionName As String = "ICustomerDO_CreateContactHistory"
    

    Dim xmlClassDefDoc As FreeThreadedDOMDocument40
    Dim objIClassDef As IomCustClassDef
    Dim xmlDateTime As IXMLDOMElement
    Dim xmlTableElem As IXMLDOMElement
    Dim xmlCustNo As IXMLDOMElement
    Dim xmlOut As FreeThreadedDOMDocument40
                
    Set xmlDateTime = m_objXmlAssist.GetNode(vxmlTableElement, ".//CONTACTHISTORYDATETIME", False)
                              
    If xmlDateTime Is Nothing Then
        Set xmlDateTime = vxmlTableElement.ownerDocument.createElement("CONTACTHISTORYDATETIME")
        vxmlTableElement.appendChild xmlDateTime
    End If
    
    If Len(xmlDateTime.Text) = 0 Then
        xmlDateTime.Text = Format(Now, "dd/mm/yyyy hh:mm:ss")
    End If
    
    Set objIClassDef = New omCustClassDef
    Set xmlClassDefDoc = objIClassDef.LoadContactHistoryData()
    m_objIDOAssist.Create vxmlTableElement, xmlClassDefDoc
    
    Set xmlOut = New FreeThreadedDOMDocument40
    xmlOut.validateOnParse = False
    xmlOut.setProperty "NewParser", True
    
    Set xmlTableElem = xmlOut.createElement("CONTACTHISTORY")
    xmlOut.appendChild xmlTableElem
    
    Set xmlCustNo = m_objXmlAssist.GetNode(vxmlTableElement, ".//CUSTOMERNUMBER")
    xmlTableElem.appendChild xmlCustNo.cloneNode(True)
           
    xmlTableElem.appendChild xmlDateTime.cloneNode(True)
           
    Set ICustomerDO_CreateContactHistory = xmlTableElem
    
    Set objIClassDef = Nothing
    Set xmlClassDefDoc = Nothing
    Set xmlDateTime = Nothing
    Set xmlTableElem = Nothing
    Set xmlCustNo = Nothing
    Set xmlOut = Nothing
    
    m_objContext.SetComplete


    Exit Function

CreateContactHistoryVbErr:
    

    Set objIClassDef = Nothing
    Set xmlClassDefDoc = Nothing
    
    m_objErrAssist.AddToErrSource (strFunctionName)
    
    If m_objErrAssist.IsSystemError = True Then
        m_objErrAssist.LogError TypeName(Me), strFunctionName, Err.Number, Err.Description
    End If
    
    If m_objErrAssist.IsSystemError Then
        m_objContext.SetAbort
    Else
        m_objContext.SetComplete
    End If
        
    '   re-raise error for business object to interpret as appropriate
    Err.Raise Err.Number, Err.Source, Err.Description
    
End Function

Private Sub ICustomerDO_CreateCurrentProperty(ByVal vxmlTableElement As IXMLDOMElement)
' header ----------------------------------------------------------------------------------
' description:
'   create an instance of the persistant data associated with this data object
'   for each set of data in the request on the CurrentProperty Table
' pass:
'   vxmlTableElement  xml element containing the request
' return:                n/a
'------------------------------------------------------------------------------------------
On Error GoTo CreateCurrentPropertyVbErr

    Const strFunctionName As String = "ICustomerDO_CreateCurrentProperty"
   
    Dim xmlClassDefDoc As FreeThreadedDOMDocument40
    Dim objIClassDef As IomCustClassDef
    
                               
    Set objIClassDef = New omCustClassDef
    Set xmlClassDefDoc = objIClassDef.LoadCurrentPropertyData()
    m_objIDOAssist.Create vxmlTableElement, xmlClassDefDoc
    
    Set objIClassDef = Nothing
    Set xmlClassDefDoc = Nothing

    If Not m_objContext Is Nothing Then
        m_objContext.SetComplete
    End If

    Exit Sub

CreateCurrentPropertyVbErr:
    
    Set objIClassDef = Nothing
    Set xmlClassDefDoc = Nothing
    
    m_objErrAssist.AddToErrSource (strFunctionName)
    
    If m_objErrAssist.IsSystemError = True Then
        m_objErrAssist.LogError TypeName(Me), strFunctionName, Err.Number, Err.Description
    End If
    
    If Not m_objContext Is Nothing Then
        If m_objErrAssist.IsSystemError Then
            m_objContext.SetAbort
        Else
            m_objContext.SetComplete
        End If
    End If
        
    '   re-raise error for business object to interpret as appropriate
    Err.Raise Err.Number, Err.Source, Err.Description
    
End Sub

Private Sub ICustomerDO_CreateHomeInsurance(ByVal vxmlTableElement As IXMLDOMElement)
' header ----------------------------------------------------------------------------------
' description:
'   create an instance of the persistant data associated with this data object
'   for each set of data in the request in the HomeInsurance Table.
' pass:
'   vxmlTableElement  xml element containing the request
' return:                n/a
'------------------------------------------------------------------------------------------
On Error GoTo CreateHomeInsuranceVbErr

    Const strFunctionName As String = "ICustomerDO_CreateHomeInsurance"
   
    Dim xmlClassDefDoc As FreeThreadedDOMDocument40
    Dim objIClassDef As IomCustClassDef
    
                             
    Set objIClassDef = New omCustClassDef
    Set xmlClassDefDoc = objIClassDef.LoadHomeInsuranceData()
    m_objIDOAssist.Create vxmlTableElement, xmlClassDefDoc
    
    Set objIClassDef = Nothing
    Set xmlClassDefDoc = Nothing

    If Not m_objContext Is Nothing Then
        m_objContext.SetComplete
    End If

    Exit Sub

CreateHomeInsuranceVbErr:
    
    Set objIClassDef = Nothing
    Set xmlClassDefDoc = Nothing
    
    m_objErrAssist.AddToErrSource (strFunctionName)
    
    If m_objErrAssist.IsSystemError = True Then
        m_objErrAssist.LogError TypeName(Me), strFunctionName, Err.Number, Err.Description
    End If
    
    If Not m_objContext Is Nothing Then
        If m_objErrAssist.IsSystemError Then
            m_objContext.SetAbort
        Else
            m_objContext.SetComplete
        End If
    End If
        
    '   re-raise error for business object to interpret as appropriate
    Err.Raise Err.Number, Err.Source, Err.Description
    
End Sub

Private Sub ICustomerDO_CreatePropertyInsuranceDetails(ByVal vxmlTableElement As IXMLDOMElement)
' header ----------------------------------------------------------------------------------
' description:
'   create an instance of the persistant data associated with this data object
'   for each set of data in the request
' pass:
'   vxmlTableElement  xml element containing the request
' return:                n/a
'------------------------------------------------------------------------------------------
On Error GoTo CreatePropertyInsuranceDetailsVbErr

    Const strFunctionName As String = "ICustomerDO_CreatePropertyInsuranceDetails"
   
    Dim xmlClassDefDoc As FreeThreadedDOMDocument40
   
    Dim xmlTempResponseNode As IXMLDOMNode
    
    ICustomerDO_CreateCurrentProperty vxmlTableElement
    ICustomerDO_CreateHomeInsurance vxmlTableElement
    

    If Not m_objContext Is Nothing Then
        m_objContext.SetComplete
    End If

    Exit Sub

CreatePropertyInsuranceDetailsVbErr:
    
    
    m_objErrAssist.AddToErrSource (strFunctionName)
    
    If m_objErrAssist.IsSystemError = True Then
        m_objErrAssist.LogError TypeName(Me), strFunctionName, Err.Number, Err.Description
    End If
    
    If Not m_objContext Is Nothing Then
        If m_objErrAssist.IsSystemError Then
            m_objContext.SetAbort
        Else
            m_objContext.SetComplete
        End If
    End If
        
    '   re-raise error for business object to interpret as appropriate
    Err.Raise Err.Number, Err.Source, Err.Description
    
End Sub

Private Function ICustomerDO_CreateAlias(ByVal vxmlTableElement As IXMLDOMElement) As IXMLDOMElement
' header ------------------------------------------------------------------------
' description:
'   create an instance of the persistant data associated with tables 'Person'
'   and 'Alias' for each set of data in the request
' pass:
'   vstrXMLRequest  xml Request data stream containing data to be persisted
' Dependencies : CustomerBO, CustomerTxBO
' Raise Errors:
'-------------------------------------------------------------------------------
On Error GoTo ICustomerDO_CreateAliasVbErr

    Const strFunctionName As String = "ICustomerDO_CreateAlias"

    Dim objContext As ObjectContext
    Set objContext = GetObjectContext()

    Dim strTemp As String, strPersonGuid As String
    Dim strTagName As String
    Dim blnIsPersonDataAffected As Boolean 'any changes to be made to the Person Date
    Dim blnIsNewPerson As Boolean 'should this create a new person data or update an existing one

    Dim objXmlAssist As New XMLAssist
    Dim objErrAssist As New ErrAssist
    Dim objIDOAssist As IDOAssist
    Dim objIPersonDO As IPersonDO
    Dim objIClassDef As IomCustClassDef
    
    Dim xmlResponseElem As IXMLDOMElement
    Dim xmlElement As IXMLDOMElement
    Dim xmlAliasClassDef As New FreeThreadedDOMDocument40
    xmlAliasClassDef.validateOnParse = False
    xmlAliasClassDef.setProperty "NewParser", True

    If objContext Is Nothing Then
        Set objIDOAssist = New DOAssist
        Set objIPersonDO = New PersonDO
    Else
        Set objIDOAssist = objContext.CreateInstance(gstrBASE_COMPONENT & ".DOAssist")
        Set objIPersonDO = objContext.CreateInstance(cstrPERSON_COMPONENT & ".PersonDO")
    End If

    Set xmlElement = vxmlTableElement.getElementsByTagName("PERSON").Item(0)
    If xmlElement Is Nothing Then
        blnIsPersonDataAffected = False
    Else
        blnIsPersonDataAffected = True
        strTemp = objXmlAssist.GetTagValue(xmlElement, "PERSONGUID")
        blnIsNewPerson = IIf(Len(strTemp) = 0, True, False)
    End If

    'Create a new or modify existing record in the table 'PERSON'
    If blnIsPersonDataAffected Then
        If blnIsNewPerson Then 'Create new person data.
            Set xmlElement = objIPersonDO.Create(xmlElement)
            ' Assign PersonGUID to Alias node
            strPersonGuid = xmlElement.getElementsByTagName("PERSONGUID").Item(0).Text
            Set xmlElement = vxmlTableElement.getElementsByTagName("PERSONGUID").Item(0)
            If xmlElement Is Nothing Then
                Set xmlElement = vxmlTableElement.ownerDocument.createElement("PERSONGUID")
                vxmlTableElement.appendChild xmlElement
            End If
            xmlElement.Text = strPersonGuid
        Else 'Update the person data
            Call objIPersonDO.Update(xmlElement)
        End If
    End If

    ' Find new Alias Sequence Number and assign to Input and Output
    Set objIClassDef = New omCustClassDef
    Set xmlAliasClassDef = objIClassDef.LoadAliasData

    Call objIDOAssist.GenerateSequenceNumber(vxmlTableElement, xmlAliasClassDef, _
                                               "ALIASSEQUENCENUMBER")
    strTemp = vxmlTableElement.getElementsByTagName("ALIASSEQUENCENUMBER").Item(0).Text
    
    Set xmlResponseElem = vxmlTableElement.ownerDocument.createElement("RESPONSE")
    Set xmlElement = vxmlTableElement.ownerDocument.createElement("ALIASSEQUENCENUMBER")
    xmlElement.Text = strTemp
    xmlResponseElem.appendChild xmlElement
    
    ' Add the respective CustomerNumber and VersionNumber along with the Sequence numner
    Dim strCustomerNumber As String, strCustomerVersionNumber As String
    strCustomerNumber = vxmlTableElement.getElementsByTagName("CUSTOMERNUMBER").Item(0).Text
    strCustomerVersionNumber = _
      vxmlTableElement.getElementsByTagName("CUSTOMERVERSIONNUMBER").Item(0).Text
    
    Set xmlElement = vxmlTableElement.ownerDocument.createElement("ALIASSEQUENCENUMBER")
    xmlElement.Text = strTemp
    xmlResponseElem.appendChild xmlElement
    
    ' Create Alias Record
    objIDOAssist.Create vxmlTableElement, xmlAliasClassDef
    
    Set ICustomerDO_CreateAlias = xmlResponseElem
    
    Set objXmlAssist = Nothing
    Set objErrAssist = Nothing
    Set objIDOAssist = Nothing
    Set objIPersonDO = Nothing
    
    Set xmlResponseElem = Nothing
    Set xmlElement = Nothing
    Set xmlAliasClassDef = Nothing

    If Not objContext Is Nothing Then
        objContext.SetComplete
        Set objContext = Nothing
    End If

    Exit Function
ICustomerDO_CreateAliasVbErr:
    Set objXmlAssist = Nothing
    Set objIDOAssist = Nothing
    Set objIPersonDO = Nothing

    Set xmlResponseElem = Nothing
    Set xmlElement = Nothing
    Set xmlAliasClassDef = Nothing

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
        Set objContext = Nothing
    End If

    Set objErrAssist = Nothing

    '  re-raise error for business object to interpret as appropriate
    Err.Raise Err.Number, Err.Source, Err.Description
End Function

Private Function ICustomerDO_CreateCustomerAddress( _
                    ByVal vXMLRequest As IXMLDOMElement) As IXMLDOMElement
' header ------------------------------------------------------------------------
' description:
'   create an instance of the persistant data associated with tables 'Customer Address',
'  'Address' and 'Current Property' for each set of data in the request
' pass:
'   vxmlRequest       :  xml Request Element containing data to be persisted
'   vxmlGeneratedKeys :  xml Node to append the results to
' Dependencies : CustomerBO, CustomerTxBO
' Raise Errors:
'-------------------------------------------------------------------------------
On Error GoTo CreateCustomerAddressVbErr

    Const strFunctionName As String = "ICustomerDO_CreateCustomerAddress"

    Dim strTemp As String
    Dim strTagName As String
    Dim strAddressGuid As String

    Dim blnIsAddressDataAffected As Boolean 'any changes to be made to the Addess Data
    Dim intNewSequenceNumber As Integer   'sequence number to be assigned to Alias

    Dim objXmlAssist As New XMLAssist
    Dim objErrAssist As New ErrAssist
    Dim objIDOAssist As IDOAssist
    Dim objIAddressBO As IAddressBO
    Dim objIClassDef As IomCustClassDef

    Dim xmlCustAddressGuidElem As IXMLDOMElement 'Address GUID in CustomerAddress data
    Dim xmlCustAddressOutElem As IXMLDOMElement
    Dim xmlAddressOutElem As IXMLDOMElement
    Dim xmlSeqNumberElem As IXMLDOMElement
    Dim xmlElement As IXMLDOMElement
    Dim xmlCustomerAddressDef As FreeThreadedDOMDocument40

    Dim intOperationOnAddress As Integer
    
    Dim objContext As ObjectContext
    Set objContext = GetObjectContext()

    If objContext Is Nothing Then
        Set objIDOAssist = New DOAssist
        Set objIAddressBO = New AddressBO
    Else
        Set objIDOAssist = objContext.CreateInstance(gstrBASE_COMPONENT & ".DOAssist")
        Set objIAddressBO = objContext.CreateInstance(cstrADDRESS_COMPONENT & ".AddressBO")
    End If

    'Find out whether any chages are to be made to the table 'Address
    blnIsAddressDataAffected = IIf(objXmlAssist.GetTagValue(vXMLRequest, "ADDRESS") _
                                        <> "", True, False)
    
    Set xmlCustAddressGuidElem = vXMLRequest.selectSingleNode("ADDRESSGUID")

    'Create a new or modify existing record in the table 'ADDRESS' and assign AddressGuid to
    ' to appropriate elements
    If blnIsAddressDataAffected Then
        Set xmlElement = vXMLRequest.getElementsByTagName("ADDRESS").Item(0)
        Set xmlElement = objIAddressBO.SaveAddress(xmlElement, intOperationOnAddress)  ' Save the Address
        objErrAssist.CheckXMLResponse xmlElement.xml, True

        If intOperationOnAddress = booCreate Then
            strAddressGuid = objXmlAssist.GetTagValue(xmlElement, "ADDRESSGUID")
            
            ' Assign the AddressGuid to CustomerAddress data
            xmlCustAddressGuidElem.Text = strAddressGuid
            
        Else 'Get AddressGuid from 'ADDRESS' in xmlIn
            strAddressGuid = objXmlAssist.GetTagValue(xmlElement, "ADDRESSGUID")
        End If

    Else 'Address Guid must be mentioned in the 'CUSTOMERADDRESS' tag of Input
        strAddressGuid = xmlCustAddressGuidElem.Text
        If strAddressGuid = "" Then
            objErrAssist.ThrowError TypeName(Me), strFunctionName, 507, _
                        "ADDRESSGUID" & " not found for CUSTOMERADDRESS"
        End If
    End If
    
    Set objIClassDef = New omCustClassDef
    Set xmlCustomerAddressDef = objIClassDef.LoadCustomerAddressData
    Call objIDOAssist.GenerateSequenceNumber(vXMLRequest, xmlCustomerAddressDef, _
                                                "CUSTOMERADDRESSSEQUENCENUMBER")

    ' Add the CustomerAddress Keys and AddressGuID to GeneratedKeys node
    Set xmlCustAddressOutElem = vXMLRequest.ownerDocument.createElement("CUSTOMERADDRESS")

    Set xmlElement = vXMLRequest.ownerDocument.createElement("CUSTOMERNUMBER")
    xmlElement.Text = objXmlAssist.GetMandatoryElementText(vXMLRequest, "CUSTOMERNUMBER")
    xmlCustAddressOutElem.appendChild xmlElement

    Set xmlElement = vXMLRequest.ownerDocument.createElement("CUSTOMERVERSIONNUMBER")
    xmlElement.Text = objXmlAssist.GetMandatoryElementText(vXMLRequest, "CUSTOMERVERSIONNUMBER")
    xmlCustAddressOutElem.appendChild xmlElement

    Set xmlElement = vXMLRequest.ownerDocument.createElement("CUSTOMERADDRESSSEQUENCENUMBER")
    xmlElement.Text = objXmlAssist.GetMandatoryElementText(vXMLRequest, _
                                            "CUSTOMERADDRESSSEQUENCENUMBER")
    xmlCustAddressOutElem.appendChild xmlElement

    ' Add the AddressGUID only when a new address was created
    If intOperationOnAddress = booCreate Then
        Set xmlAddressOutElem = vXMLRequest.ownerDocument.createElement("ADDRESS")
        xmlCustAddressOutElem.appendChild xmlAddressOutElem
        
        Set xmlElement = vXMLRequest.ownerDocument.createElement("ADDRESSSGUID")
        xmlElement.Text = strAddressGuid
        xmlAddressOutElem.appendChild xmlElement
    End If
    
    ' create new record in the table 'CustomerAddress'
    Call objIDOAssist.Create(vXMLRequest, xmlCustomerAddressDef)

    'Create Current Property record, if required. Assign Sequence number before creation
    Set xmlElement = vXMLRequest.selectSingleNode("CURRENTPROPERTY")
    If Not xmlElement Is Nothing Then
        Set xmlSeqNumberElem = xmlElement.selectSingleNode("CUSTOMERADDRESSSEQUENCENUMBER")
        If xmlSeqNumberElem Is Nothing Then
            Set xmlSeqNumberElem = vXMLRequest.ownerDocument.createElement("CUSTOMERADDRESSSEQUENCENUMBER")
            xmlElement.appendChild xmlSeqNumberElem
        End If
        xmlSeqNumberElem.Text = vXMLRequest.selectSingleNode("CUSTOMERADDRESSSEQUENCENUMBER").Text
        objIDOAssist.Create xmlElement, objIClassDef.LoadCurrentPropertyData
    End If

    Set ICustomerDO_CreateCustomerAddress = xmlCustAddressOutElem.cloneNode(True)
    
    Set objXmlAssist = Nothing
    Set objErrAssist = Nothing
    Set objIDOAssist = Nothing
    Set objIAddressBO = Nothing

    Set xmlCustAddressGuidElem = Nothing
    Set xmlCustAddressOutElem = Nothing
    Set xmlElement = Nothing
    Set xmlSeqNumberElem = Nothing

    If Not objContext Is Nothing Then
        objContext.SetComplete
        Set objContext = Nothing
    End If

    Exit Function
CreateCustomerAddressVbErr:
    Set objXmlAssist = Nothing
    Set objErrAssist = Nothing
    Set objIDOAssist = Nothing
    Set objIAddressBO = Nothing

    Set xmlCustAddressGuidElem = Nothing
    Set xmlCustAddressOutElem = Nothing
    Set xmlElement = Nothing
    Set xmlSeqNumberElem = Nothing

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
        Set objContext = Nothing
    End If

    Set objErrAssist = Nothing

    '  re-raise error for business object to interpret as appropriate
    Err.Raise Err.Number, Err.Source, Err.Description
End Function

Private Function ICustomerDO_CreateDependant(ByVal vxmlTableElement As IXMLDOMElement) _
                                                As IXMLDOMElement
' header ----------------------------------------------------------------------------------
' description:
'   create an instance of the persistant data associated with this data object
'   for each set of data in the request
' pass:
'   vXMLRequest  xml Request node containing data to be persisted
' return :
'       node with the generated keys
' Raise Errors:
'------------------------------------------------------------------------------------------
On Error GoTo ICustomerDO_CreateDependantVbErr

    Const strFunctionName As String = "ICustomerDO_CreateDependant"

    Dim objContext As ObjectContext
    Set objContext = GetObjectContext()

    Dim objXmlAssist As New XMLAssist
    Dim xmlElement As IXMLDOMElement
    Dim objErrAssist As New ErrAssist

    Dim objIDOAssist As IDOAssist
    Dim objIClassDef As IomCustClassDef
    Dim xmlClassDef As FreeThreadedDOMDocument40

    Dim strPersonGuid As String
    
    If objContext Is Nothing Then
        Set objIDOAssist = New DOAssist
    Else
        Set objIDOAssist = objContext.CreateInstance(gstrBASE_COMPONENT & ".DOAssist")
    End If
    
    ' If Person data is mentioned in Request, save it
    Dim intOperationOnPersion As Integer
    If Not vxmlTableElement.selectSingleNode("//PERSON") Is Nothing Then
        Call SavePerson(vxmlTableElement, strPersonGuid, intOperationOnPersion)

        Set xmlElement = vxmlTableElement.selectSingleNode("PERSONGUID")
        If xmlElement Is Nothing Then
            Set xmlElement = vxmlTableElement.appendChild(vxmlTableElement.ownerDocument.createElement("PERSONGUID"))
        End If
        xmlElement.Text = strPersonGuid
    Else
        Set xmlElement = vxmlTableElement.selectSingleNode(".//PERSONGUID")

        ' if PersonGUID is null in Dependant data, raise error
        If xmlElement Is Nothing Then
            objErrAssist.ThrowError TypeName(Me), strFunctionName, oeInvalidParameter, _
                            " PERSONGUID is not given in Dependant data "
        ElseIf xmlElement.Text = "" Then
            objErrAssist.ThrowError TypeName(Me), strFunctionName, oeInvalidParameter, _
                            " PERSONGUID is not given in Dependant data "
        End If
    End If
    
    '
    ' Insert the dependant
    '
    Set objIClassDef = New omCustClassDef
    Set xmlClassDef = objIClassDef.LoadDependantData
    objIDOAssist.GenerateSequenceNumber vxmlTableElement, xmlClassDef, "DEPENDANTSEQUENCENUMBER"
    objIDOAssist.Create vxmlTableElement, xmlClassDef
    
    Set xmlElement = vxmlTableElement.ownerDocument.createElement("DEPENDANTSEQUENCENUMBER")
    xmlElement.Text = objXmlAssist.GetTagValue(vxmlTableElement, "DEPENDANTSEQUENCENUMBER")
    
    Set ICustomerDO_CreateDependant = xmlElement
    
    Set objIDOAssist = Nothing
    Set objErrAssist = Nothing
    Set xmlElement = Nothing
    Set objXmlAssist = Nothing
    Set xmlClassDef = Nothing
    Set objIClassDef = Nothing
    If Not objContext Is Nothing Then
        objContext.SetComplete
        Set objContext = Nothing
    End If

    Exit Function

ICustomerDO_CreateDependantVbErr:

    Set objIDOAssist = Nothing
    Set xmlElement = Nothing
    Set objXmlAssist = Nothing
    Set xmlClassDef = Nothing
    Set objIClassDef = Nothing
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
    
End Function

Private Sub ICustomerDO_DeleteAlias(ByVal vxmlTableElement As IXMLDOMElement)
' header ----------------------------------------------------------------------------------
' description:
'   Delete a single instance of the persistant data associated with table 'Alias'.
'   And also delete the respective record from table 'PERSON', subject to that the
'   particular record is not referred to by any other data than the current Alias.
' pass:
'   vxmlTableElem   xml Request node which identifies the instance to be deleted (Alias)
' Raise Errors:
'------------------------------------------------------------------------------------------
On Error GoTo ICustomerDO_DeleteAliasVbErr

    Const strFunctionName As String = "ICustomerDO_DeleteAlias"

    Dim objErrAssist As New ErrAssist

    Dim objContext As ObjectContext
    Set objContext = GetObjectContext()

    Dim objIDOAssist As IDOAssist
    Dim objIClassDef As IomCustClassDef

    If objContext Is Nothing Then
        Set objIDOAssist = New DOAssist
    Else
        Set objIDOAssist = objContext.CreateInstance(gstrBASE_COMPONENT & ".DOAssist")
    End If
    
    Set objIClassDef = New omCustClassDef
    objIDOAssist.Delete vxmlTableElement, objIClassDef.LoadAliasData

    Call DeletePersonForThisAlias(vxmlTableElement) 'delete the respective record from Person

    Set objIDOAssist = Nothing
    Set objIClassDef = Nothing
    Set objErrAssist = Nothing

    If Not objContext Is Nothing Then
        objContext.SetComplete
        Set objContext = Nothing
    End If

    Exit Sub

ICustomerDO_DeleteAliasVbErr:

    Set objIDOAssist = Nothing
    Set objIClassDef = Nothing
    
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
         Set objContext = Nothing
    End If

    Set objErrAssist = Nothing
    '   re-raise error for business object to interpret as appropriate
    Err.Raise Err.Number, Err.Source, Err.Description

End Sub

Private Function ICustomerDO_FindContactHistoryList(ByVal vxmlTableElement As IXMLDOMElement) As IXMLDOMNode
    
    On Error GoTo FindContactHistoryListVbErr

    Const strFunctionName As String = "ICustomerDO_FindContactHistoryList"
    

    If m_objContext Is Nothing Then
        m_objErrAssist.ThrowError TypeName(Me), strFunctionName, oeMTSNotFound
    End If
        
    Dim xmlClassDefDoc As FreeThreadedDOMDocument40
    Dim objIClassDef As IomCustClassDef
    Dim xmlSortXML As IXMLDOMNode
    
    Dim strCustomerNumber As String
    Dim strSortOrder As String
    Dim strField As String
    
    strCustomerNumber = m_objXmlAssist.GetNodeValue(vxmlTableElement, ".//CUSTOMERNUMBER")
    strSortOrder = m_objXmlAssist.GetNodeValue(vxmlTableElement, ".//SORTORDER")
              
    Select Case strSortOrder
        
        Case "1"
            strField = "CONTACTREASONCODE"
        Case "2"
            'MV - 09/05/02 - BMIDS00004
            strField = "CONTACTHISTORYDATETIME DESC "
        Case "3"
            strField = "USERID"
        Case Else
            m_objErrAssist.ThrowError TypeName(Me), _
                                      strFunctionName, _
                                      oeInvalidParameter, _
                                      "Sort order must be 1,2 or 3"
    End Select
     
    Set objIClassDef = New omCustClassDef
    Set xmlClassDefDoc = objIClassDef.LoadContactHistoryData()
    Set ICustomerDO_FindContactHistoryList = m_objIDOAssist.FindList(vxmlTableElement, xmlClassDefDoc, strField)
   
    Set objIClassDef = Nothing
    Set xmlClassDefDoc = Nothing
    
    m_objContext.SetComplete
    
    
    Exit Function

FindContactHistoryListVbErr:
    
    
    Set objIClassDef = Nothing
    Set xmlClassDefDoc = Nothing
    
    m_objErrAssist.AddToErrSource (strFunctionName)
    
    If m_objErrAssist.IsSystemError = True Then
        m_objErrAssist.LogError TypeName(Me), strFunctionName, Err.Number, Err.Description
    End If
    
    If m_objErrAssist.IsSystemError Then
        m_objContext.SetAbort
    Else
        m_objContext.SetComplete
    End If
    
    '   re-raise error for business object to interpret as appropriate
    Err.Raise Err.Number, Err.Source, Err.Description

End Function


Private Function ICustomerDO_GetComponentData( _
                    ByVal vXMLRequest As IXMLDOMElement) As IXMLDOMElement
' header ----------------------------------------------------------------------------------
' description:
'   Get the specifed data for a single instance of the persistant data associated with
'   this data object
' pass:
'   vXMLRequest  xml Request node containing data to which identifies
'                   the instance of the persistant data to be retrieved
' return:
'               node containing XML data retrieved
'------------------------------------------------------------------------------------------
On Error GoTo ICustomerDO_GetComponentDataVbErr

    Const strFunctionName As String = "ICustomerDO_GetComponentData"

    Dim objErrAssist As New ErrAssist

    Dim objIDOAssist As IDOAssist
    Dim objIClassDef As IomCustClassDef
    
    Dim xmlResponseElem As IXMLDOMElement
    Dim xmlClassDef As FreeThreadedDOMDocument40
    Dim xmlTableNode As IXMLDOMNode
    Dim objXmlAssist As New XMLAssist
    
    Dim strTableName As String
      
    Dim objContext As ObjectContext
    Set objContext = GetObjectContext()

    If objContext Is Nothing Then
        Set objIDOAssist = New DOAssist
    Else
        Set objIDOAssist = objContext.CreateInstance(gstrBASE_COMPONENT & ".DOAssist")
    End If

    Set xmlResponseElem = vXMLRequest.ownerDocument.createElement("RESPONSE")
    xmlResponseElem.setAttribute "TYPE", "SUCCESS"

    '   Load the Class defintion for the table
    If vXMLRequest.nodeName = "TABLE" Then
        Set xmlTableNode = vXMLRequest
    Else
        Set xmlTableNode = vXMLRequest.selectSingleNode(".//TABLE")
    End If
    strTableName = xmlTableNode.firstChild.Text
    Set objIClassDef = New omCustClassDef
    Set xmlClassDef = objIClassDef.LoadDefinition(strTableName)
       
    Set ICustomerDO_GetComponentData = objIDOAssist.GetComponentData(vXMLRequest, xmlClassDef)

    Set objIDOAssist = Nothing
    Set objErrAssist = Nothing

    If Not objContext Is Nothing Then
        objContext.SetComplete
    End If

    Set xmlResponseElem = Nothing
    Set xmlClassDef = Nothing
    Set objXmlAssist = Nothing
    Set objErrAssist = Nothing

    Exit Function

ICustomerDO_GetComponentDataVbErr:
    Set xmlResponseElem = Nothing
    Set objXmlAssist = Nothing

    Set objIDOAssist = Nothing

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

End Function

Private Function ICustomerDO_GetContactHistory(ByVal vxmlTableElement As IXMLDOMElement) As IXMLDOMNode
' header ----------------------------------------------------------------------------------
' description:
'   Get the data for a single instance of the persistant data associated with
'   this data object
' pass:
'   vxmlTableElement  xml element containing the request
' return:                xml node containing retrieved data
'------------------------------------------------------------------------------------------
On Error GoTo GetContactHistoryVbErr

    Const strFunctionName As String = "ICustomerDO_GetContactHistory"
    
    
    If m_objContext Is Nothing Then
        m_objErrAssist.ThrowError TypeName(Me), strFunctionName, oeMTSNotFound
    End If
                   
    Dim xmlClassDefDoc As FreeThreadedDOMDocument40
    Dim objIClassDef As IomCustClassDef
    
    Set objIClassDef = New omCustClassDef
    Set xmlClassDefDoc = objIClassDef.LoadContactHistoryData()
    Set ICustomerDO_GetContactHistory = m_objIDOAssist.GetData(vxmlTableElement, xmlClassDefDoc)

    Set objIClassDef = Nothing
    Set xmlClassDefDoc = Nothing
    
    m_objContext.SetComplete
    
    
    Exit Function

GetContactHistoryVbErr:
    
    
    Set objIClassDef = Nothing
    Set xmlClassDefDoc = Nothing
    
    m_objErrAssist.AddToErrSource (strFunctionName)
    
    If m_objErrAssist.IsSystemError = True Then
        m_objErrAssist.LogError TypeName(Me), strFunctionName, Err.Number, Err.Description
    End If
    
    If m_objErrAssist.IsSystemError Then
        m_objContext.SetAbort
    Else
        m_objContext.SetComplete
    End If
        
    '   re-raise error for business object to interpret as appropriate
    Err.Raise Err.Number, Err.Source, Err.Description
    
End Function

Private Function ICustomerDO_GetCustomerFromLegacyNumber(ByVal vXMLRequest As IXMLDOMElement) As IXMLDOMNode
' header ----------------------------------------------------------------------------------
' description:
'   Get the data for all instances of the persistant data associated with
'   this data object for the values supplied
' pass:
'   vxmlTableElement  xml element containing the request
' return:                xml node containing retrieved data
'------------------------------------------------------------------------------------------
On Error GoTo GetCustomerFromLegacyNumberVbErr

    Const strFunctionName As String = "ICustomerDO_GetCustomerFromLegacyNumber"
    
        
    Dim xmlCustomerClassDefDoc As FreeThreadedDOMDocument40
    Dim xmlCustVersClassDefDoc As FreeThreadedDOMDocument40
    Dim xmlOut As FreeThreadedDOMDocument40
    Dim objIClassDef As IomCustClassDef
    'Dim objIADOAssist As IADOAssist
    Dim xmlCustomerElem As IXMLDOMElement
    
    Dim adoConnection As ADODB.Connection
    Dim adoCommand As ADODB.Command
    Dim adoRecordset As ADODB.Recordset
    'Dim adoParameter As ADODB.Parameter
    
    Dim strOtherSystemCustomerNumber As String
    
    ' PSC 03/10/03 BMIDS612
    Dim strCustomerNumber As String
        
    ' Check the correct keys have been passed in
    strOtherSystemCustomerNumber = m_objXmlAssist.GetNodeValue(vXMLRequest, _
                                                             ".//OTHERSYSTEMCUSTOMERNUMBER")
    
    ' PSC 03/10/03 BMIDS612
    strCustomerNumber = m_objXmlAssist.GetNodeValue(vXMLRequest, ".//CUSTOMERNUMBER", False)
    ' PSC 03/10/03 BMIDS612
    
    'Set objIADOAssist = New ADOAssist
    
    Set adoConnection = New ADODB.Connection
    Set adoCommand = New ADODB.Command
    Set adoRecordset = New ADODB.Recordset
        
    'EP974/MAR1848
    'Dim strSQL As String
    '
    '' PSC 04/10/03 BMIDS612
    '
    'strSQL = "SELECT * FROM CUSTOMER C, CUSTOMERVERSION CV, " & _
    '         "(SELECT CUSTOMERNUMBER, MAX(CUSTOMERVERSIONNUMBER) MAXVERSION " & _
    '         "FROM CUSTOMERVERSION GROUP BY CUSTOMERNUMBER) VERSIONS " & _
    '         "WHERE (C.OTHERSYSTEMCUSTOMERNUMBER = ?  OR " & _
    '         "(C.OTHERSYSTEMCUSTOMERNUMBER IS NULL AND C.CUSTOMERNUMBER = ? ))" & _
    '         "AND C.CUSTOMERNUMBER = CV.CUSTOMERNUMBER " & _
    '         "AND VERSIONS.CUSTOMERNUMBER = CV.CUSTOMERNUMBER " & _
    '         "AND VERSIONS.MAXVERSION = CV.CUSTOMERVERSIONNUMBER "

    With adoCommand
        .CommandType = adCmdStoredProc
        .CommandText = "USP_GetCustomerFromLegacyNumber"
    
        ' Set up the parameters to be substituted into the SQL
        .Parameters.Append .CreateParameter("OtherSystemCustomerNumber", adBSTR, adParamInput, Len(strOtherSystemCustomerNumber), strOtherSystemCustomerNumber)
   
        ' PSC 04/10/03 BMIDS612 - Start
        .Parameters.Append .CreateParameter("CustomerNumber", adBSTR, adParamInput, Len(strCustomerNumber), strCustomerNumber)
        ' PSC 04/10/03 BMIDS612 - End
    End With
        
    ' Set Recordset properties. Must be adUseClient for disconnected recordset
    With adoRecordset
        .CursorLocation = adUseClient
        .CursorType = adOpenForwardOnly
        .LockType = adLockReadOnly
    End With
    
    With adoConnection
        .ConnectionString = adoGetDbConnectString
        .Open
    End With
    'EP974/MAR1848 End
    
    adoCommand.ActiveConnection = adoConnection
    Set adoRecordset.Source = adoCommand

    'Populate recordset
    adoRecordset.Open
    
    'Disconnect the recordset
    Set adoRecordset.ActiveConnection = Nothing
    adoConnection.Close
    
    If adoRecordset.EOF Then
        m_objErrAssist.ThrowError TypeName(Me), strFunctionName, oeRecordNotFound
    End If
     
    Set objIClassDef = New omCustClassDef
    Set xmlCustomerClassDefDoc = objIClassDef.LoadCustomerData()
    Set xmlCustVersClassDefDoc = objIClassDef.LoadCustomerVersionData()
    
    Set xmlOut = New FreeThreadedDOMDocument40
    xmlOut.validateOnParse = False
    xmlOut.setProperty "NewParser", True
    
    ' Convert recordset to XML
    Set xmlCustomerElem = m_objIDOAssist.GetXMLFromRecordSet(adoRecordset, xmlCustomerClassDefDoc)
    
    m_objIDOAssist.GetXMLFromRecordSet adoRecordset, xmlCustVersClassDefDoc, xmlCustomerElem
    
    Set ICustomerDO_GetCustomerFromLegacyNumber = xmlCustomerElem

    adoRecordset.Close
    
    
    Set objIClassDef = Nothing
    Set xmlCustomerClassDefDoc = Nothing
    Set xmlCustVersClassDefDoc = Nothing
    Set xmlOut = Nothing
    'Set objIADOAssist = Nothing
    Set adoConnection = Nothing
    Set adoCommand = Nothing
    Set adoRecordset = Nothing
    'Set adoParameter = Nothing
    Set xmlCustomerElem = Nothing
    
    m_objContext.SetComplete
    
    
    Exit Function

GetCustomerFromLegacyNumberVbErr:
    
    
    ' Close the recordset if still open
    If Not adoRecordset Is Nothing Then
        If adoRecordset.State = adStateOpen Then
            adoRecordset.Close
        End If
    End If
    
    ' Close the connection if still open
    If Not adoConnection Is Nothing Then
        If adoConnection.State = adStateOpen Then
            adoConnection.Close
        End If
    End If
    
    Set objIClassDef = Nothing
    Set xmlCustomerClassDefDoc = Nothing
    Set xmlCustVersClassDefDoc = Nothing
    Set xmlOut = Nothing
    'Set objIADOAssist = Nothing
    Set adoConnection = Nothing
    Set adoCommand = Nothing
    Set adoRecordset = Nothing
    'Set adoParameter = Nothing
    Set xmlCustomerElem = Nothing
    
    m_objErrAssist.AddToErrSource (strFunctionName)
    
    If m_objErrAssist.IsSystemError = True Then
        m_objErrAssist.LogError TypeName(Me), strFunctionName, Err.Number, Err.Description
    End If
    
    If m_objErrAssist.IsSystemError Then
        m_objContext.SetAbort
    Else
        m_objContext.SetComplete
    End If
    
    '   re-raise error for business object to interpret as appropriate
    Err.Raise Err.Number, Err.Source, Err.Description

End Function


Private Function ICustomerDO_GetPropertyInsuranceDetails(ByVal vXMLRequest As IXMLDOMElement) As IXMLDOMNode
' header ----------------------------------------------------------------------------------
' description:
'   Get the data for a single instance of the persistant data associated with
'   this data object
' pass:
'   vxmlRequest  xml element containing the request
' return:                xmlCurrPropNode containing retrieved data
'------------------------------------------------------------------------------------------
On Error GoTo GetPropertyInsuranceDetailsVbErr

    Const strFunctionName As String = "ICustomerDO_GetPropertyInsuranceDetails"
      
    Dim adoCommand As New ADODB.Command
    Dim objIADOAssist As IADOAssist
    Dim adoParameter As ADODB.Parameter
    Dim xmlCurrPropClassDefDoc As FreeThreadedDOMDocument40
    Dim xmlHomInsClassDefDoc As FreeThreadedDOMDocument40
    Dim objIClassDef As IomCustClassDef
    Dim xmlCurrPropNode As IXMLDOMNode
    Dim strCustomerNumber As String
    Dim strCustomerVersionNumber As String
    Dim strCustAddrSequenceNumber As String
    Dim strSQL As String
    
    If m_objContext Is Nothing Then
        Set objIADOAssist = New ADOAssist
    Else
        Set objIADOAssist = m_objContext.CreateInstance(gstrBASE_COMPONENT & ".ADOAssist")
    End If

    'Fetch the values of all the columns in the Primary Key.
    strCustomerNumber = m_objXmlAssist.GetNodeValue(vXMLRequest, ".//CUSTOMERNUMBER")
    strCustomerVersionNumber = m_objXmlAssist.GetNodeValue(vXMLRequest, ".//CUSTOMERVERSIONNUMBER")
    strCustAddrSequenceNumber = _
                    m_objXmlAssist.GetNodeValue(vXMLRequest, ".//CUSTOMERADDRESSSEQUENCENUMBER")
        
    'Build the SQL string to be executed
    strSQL = "SELECT * FROM HOMEINSURANCE, CURRENTPROPERTY " & _
             " WHERE CURRENTPROPERTY.CUSTOMERNUMBER = HOMEINSURANCE.CUSTOMERNUMBER AND" & _
             " CURRENTPROPERTY.CUSTOMERVERSIONNUMBER = HOMEINSURANCE.CUSTOMERVERSIONNUMBER AND" & _
             " CURRENTPROPERTY.CUSTOMERADDRESSSEQUENCENUMBER = HOMEINSURANCE.CUSTOMERADDRESSSEQUENCENUMBER AND" & _
             " CURRENTPROPERTY.CUSTOMERNUMBER = ? AND CURRENTPROPERTY.CUSTOMERVERSIONNUMBER = ? AND CURRENTPROPERTY.CUSTOMERADDRESSSEQUENCENUMBER = ?"
                      
    adoCommand.ActiveConnection = objIADOAssist.GetConnStr
    
    'Define the parameters
    Set adoParameter = adoCommand.CreateParameter("CustomerNumber", adBSTR, adParamInput, Len(strCustomerNumber), strCustomerNumber)
    adoCommand.Parameters.Append adoParameter
    Set adoParameter = Nothing

    Set adoParameter = adoCommand.CreateParameter("CustomerVersionNumber", adBSTR, adParamInput, Len(strCustomerVersionNumber), strCustomerVersionNumber)
    adoCommand.Parameters.Append adoParameter
    Set adoParameter = Nothing
    
    Set adoParameter = adoCommand.CreateParameter("CustomerAddSeqNo", adInteger, adParamInput)
    adoParameter.Value = strCustAddrSequenceNumber
    adoCommand.Parameters.Append adoParameter
    Set adoParameter = Nothing

    
    adoCommand.CommandText = strSQL

    Dim rstThisRecSet As New ADODB.Recordset
    Set rstThisRecSet = adoCommand.Execute(, , adCmdText)

    If rstThisRecSet.EOF And rstThisRecSet.BOF Then
        ' raise application error to be interpreted by calling object
        m_objErrAssist.ThrowError TypeName(Me), strFunctionName, oeRecordNotFound
    End If
      
    'Get xml from recordset
    Set objIClassDef = New omCustClassDef
    Set xmlCurrPropClassDefDoc = objIClassDef.LoadCurrentPropertyData()
    Set xmlHomInsClassDefDoc = objIClassDef.LoadHomeInsuranceData()
    Set xmlCurrPropNode = m_objIDOAssist.GetXMLFromRecordSet(rstThisRecSet, xmlCurrPropClassDefDoc)
    m_objIDOAssist.GetXMLFromRecordSet rstThisRecSet, xmlHomInsClassDefDoc, xmlCurrPropNode
        rstThisRecSet.Close
    Set ICustomerDO_GetPropertyInsuranceDetails = xmlCurrPropNode
    
        Set rstThisRecSet = Nothing
        Set adoCommand = Nothing
    Set objIClassDef = Nothing
    Set xmlCurrPropClassDefDoc = Nothing
    Set xmlHomInsClassDefDoc = Nothing
    Set xmlCurrPropNode = Nothing
    
    
    If Not m_objContext Is Nothing Then
        m_objContext.SetComplete
    End If
    
    Exit Function

GetPropertyInsuranceDetailsVbErr:
    
    ' Close the recordset if still open
        If Not rstThisRecSet Is Nothing Then
                If rstThisRecSet.State = adStateOpen Then
                        rstThisRecSet.Close
                End If
        End If

        Set rstThisRecSet = Nothing
        Set adoCommand = Nothing
'    Set objIClassDef = Nothing
'    Set xmlClassDefDoc = Nothing
    
    m_objErrAssist.AddToErrSource (strFunctionName)
    
    If m_objErrAssist.IsSystemError = True Then
        m_objErrAssist.LogError TypeName(Me), strFunctionName, Err.Number, Err.Description
    End If
    
    If Not m_objContext Is Nothing Then
        If m_objErrAssist.IsSystemError Then
            m_objContext.SetAbort
        Else
            m_objContext.SetComplete
        End If
    End If
        
    '   re-raise error for business object to interpret as appropriate
    Err.Raise Err.Number, Err.Source, Err.Description
    
End Function

Private Sub ICustomerDO_UpdateAlias(ByVal vxmlTableElement As IXMLDOMElement)
' header ----------------------------------------------------------------------------------
' description:
'   Update a single instance of the persistant data associated with the table 'ALIAS'
'   and the repective person data in the table 'PERSON'
' pass:
'   vXMLTableElement  xml Request data stream containing data to be persisted
' Raise Errors:
'------------------------------------------------------------------------------------------
On Error GoTo ICustomerDO_UpdateAliasVbErr

    Const strFunctionName As String = "ICustomerDO_UpdateAlias"

    Dim objContext As ObjectContext
    Set objContext = GetObjectContext()

    Dim xmlPersonGuidElem As IXMLDOMElement
    Dim xmlPersonElem As IXMLDOMElement
    Dim xmlActionOnPersonElem As IXMLDOMElement

    Dim strOperationOnPerson As String
    Dim blnPersonAffected As Boolean 'Whether person data is affected due to this request

    Dim objErrAssist As New ErrAssist
    Dim objXmlAssist As New XMLAssist

    Dim objIDOAssist As IDOAssist
    Dim objIClassDef As IomCustClassDef
    Dim objIPersonDO As IPersonDO

    If objContext Is Nothing Then
      Set objIDOAssist = New DOAssist
    Else
        Set objIDOAssist = objContext.CreateInstance(gstrBASE_COMPONENT & ".DOAssist")
    End If

    'Check whether Person data is affected due to this request
    Set xmlPersonElem = vxmlTableElement.getElementsByTagName("PERSON").Item(0)

    If xmlPersonElem Is Nothing Then
        blnPersonAffected = False
    ElseIf Len(xmlPersonElem.Text) = 0 Then
        blnPersonAffected = False
    Else
        blnPersonAffected = True
    End If

    If blnPersonAffected Then
        'Modify person based on the Input XML and determine the repective operation performed
        Set xmlActionOnPersonElem = UpdatePersonForThisAlias(xmlPersonElem)
        strOperationOnPerson = objXmlAssist.GetTagValue(xmlActionOnPersonElem, "OPERATION")
        
        'If person record was created assign this PersonGuid to current Alias.
        'If person record was deleted, assign null to PersonGuid of the the current Alias
        
        Set xmlPersonGuidElem = vxmlTableElement.getElementsByTagName("PERSONGUID").Item(0)
        If strOperationOnPerson = "CREATE" Then
            xmlPersonGuidElem.Text = xmlActionOnPersonElem.selectSingleNode("PERSONGUID").Text
        ElseIf strOperationOnPerson = "DELETE" Then
            xmlPersonGuidElem.Text = ""
        End If
    End If
    
    Set objIClassDef = New omCustClassDef
    Call objIDOAssist.Update(vxmlTableElement, objIClassDef.LoadAliasData)

    'If the operation on Person data was delete, delete the person record now
    If blnPersonAffected And strOperationOnPerson = "DELETE" Then
        If objContext Is Nothing Then
            Set objIPersonDO = New PersonDO
        Else
            Set objIPersonDO = objContext.CreateInstance(cstrPERSON_COMPONENT & ".PersonDO")
        End If
        Call objIPersonDO.Delete(xmlPersonElem)
    End If

    Set objIDOAssist = Nothing
    Set objXmlAssist = Nothing
    Set objErrAssist = Nothing
    Set objIPersonDO = Nothing

    Set xmlPersonGuidElem = Nothing
    Set xmlPersonElem = Nothing

    If Not objContext Is Nothing Then
        objContext.SetComplete
        Set objContext = Nothing
    End If

    Exit Sub

ICustomerDO_UpdateAliasVbErr:

    Set objIDOAssist = Nothing
    Set objXmlAssist = Nothing
    Set objErrAssist = Nothing
    Set objIPersonDO = Nothing

    Set xmlPersonGuidElem = Nothing
    Set xmlPersonElem = Nothing

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
        Set objContext = Nothing
    End If

    Set objErrAssist = Nothing
    '   re-raise error for business object to interpret as appropriate
    Err.Raise Err.Number, Err.Source, Err.Description

End Sub

Private Sub DeletePersonForThisAlias(ByVal vXMLRequestElem As IXMLDOMElement)
' header ----------------------------------------------------------------------------------
' description:
'   Delete a single instance of the persistant data associated with table 'Person'.
'   with PersonGuid of the current Alias Record. If there are any other references,
'   than the current alias, do not delete the Person Record and Ignore the error.
' pass: vXMLRequestElem  xml Request node which identifies the instance of Alias to be deleted
' Raise Errors:
'------------------------------------------------------------------------------------------
On Error GoTo DeletePersonForThisAliasVbErr

    Const strFunctionName As String = "DeletePersonForThisAlias"

    Dim objContext As ObjectContext
    Set objContext = GetObjectContext()

    Dim objXmlAssist As New XMLAssist
    Dim objErrAssist As New ErrAssist
    Dim xmlIn As New FreeThreadedDOMDocument40
    xmlIn.validateOnParse = False
    xmlIn.setProperty "NewParser", True

    Dim objIPersonDO As IPersonDO
    Dim strPersonGuid As String
    Dim strPersonDoInput As String

    strPersonGuid = objXmlAssist.GetTagValue(vXMLRequestElem, "PERSONGUID")
    If Len(strPersonGuid) = 0 Then
        Exit Sub
    End If

    'Build the input string to be passed to PersonDO for deletion
    strPersonDoInput = "<PERSON>"
    strPersonDoInput = strPersonDoInput & "<PERSONGUID>" & strPersonGuid & "</PERSONGUID>"
    strPersonDoInput = strPersonDoInput & "</PERSON>"
    Set xmlIn = objXmlAssist.load(strPersonDoInput, TypeName(Me), strFunctionName)
    'Delete the persistent data in PERSON table
    If Not objContext Is Nothing Then
        Set objIPersonDO = objContext.CreateInstance(cstrPERSON_COMPONENT & ".PersonDO")
    Else
        Set objIPersonDO = New PersonDO
    End If

On Error Resume Next
    Err.Clear
    objIPersonDO.Delete xmlIn.getElementsByTagName("PERSON").Item(0)
    objErrAssist.SaveErr
    If Err.Number <> 0 Then
        If objErrAssist.GetOmigaErrorNumber(Err.Number) = oeChildRecordsFound Then
            'Ignore the error
        Else
On Error GoTo DeletePersonForThisAlias2VbErr
            objErrAssist.ReRaise
        End If
    End If

On Error GoTo DeletePersonForThisAliasVbErr
    Set objXmlAssist = Nothing
    Set objErrAssist = Nothing
    Set xmlIn = Nothing

    Set objIPersonDO = Nothing

    If Not objContext Is Nothing Then
        objContext.SetComplete
    End If

    Set objContext = Nothing
    Exit Sub

DeletePersonForThisAliasVbErr:
    Set objXmlAssist = Nothing
    Set xmlIn = Nothing

    objErrAssist.AddToErrSource (strFunctionName)

    If objErrAssist.IsSystemError = True Then
        objErrAssist.LogError TypeName(Me), strFunctionName, Err.Number, Err.Description
    End If

    If Not objContext Is Nothing Then
        If objErrAssist.IsSystemError Then
            objContext.SetAbort
        ElseIf objErrAssist.GetOmigaErrorNumber(Err.Number) = oeChildRecordsFound Then
            objContext.SetAbort
        Else
            objContext.SetComplete
        End If
    End If

    Set objContext = Nothing
    Set objErrAssist = Nothing

    '  re-raise error for business object to interpret as appropriate
    Err.Raise Err.Number, Err.Source, Err.Description

    Exit Sub

DeletePersonForThisAlias2VbErr:
    Set objXmlAssist = Nothing
    Set xmlIn = Nothing

    objErrAssist.AddToErrSource (strFunctionName)

    If objErrAssist.IsSystemError = True Then
        objErrAssist.LogError TypeName(Me), strFunctionName, Err.Number, Err.Description
    End If

    If Not objContext Is Nothing Then
        objContext.SetAbort
    End If
    Set objContext = Nothing
    Set objErrAssist = Nothing

    '  re-raise error for business object to interpret as appropriate
    Err.Raise Err.Number, Err.Source, Err.Description

End Sub

Private Function UpdatePersonForThisAlias(ByVal vxmlTableElem As IXMLDOMElement) As IXMLDOMNode
' header ----------------------------------------------------------------------------------
' description:
'   Determines the operation to be performed on Person data (with respect to the RequestElem)
' pass: vxmlTableElem : xml Element containing the data to be persisted in 'Person' table
'       vxmlOutNode : Node to which the result (operation on Person, New PersonGuid, if any)
'                     are added
' Return : string giving the operation performed on Person
' Raise Errors:
'------------------------------------------------------------------------------------------

On Error GoTo UpdatePersonForThisAliasVbErr

    Const strFunctionName As String = "UpdatePersonForThisAlias"

    Dim objXmlAssist As New XMLAssist
    Dim objErrAssist As New ErrAssist

    Dim objIPersonDO As IPersonDO
    Dim objIPersonClassDef As IPersonClassDef
    Dim xmlElement As IXMLDOMElement
    Dim xmlOutPersonElem As IXMLDOMElement
    Dim xmlOperationElem As IXMLDOMElement 'Determines whether Person is created, modified or deleted

    Dim strPersonGuid As String
    Dim strPersonDoInput As String

    Dim objContext As ObjectContext
    Set objContext = GetObjectContext()

    If Not objContext Is Nothing Then
        Set objIPersonDO = objContext.CreateInstance(cstrPERSON_COMPONENT & ".PersonDO")
    Else
        Set objIPersonDO = New PersonDO
    End If

    Dim xmlDefTemp As FreeThreadedDOMDocument40

    Set xmlOutPersonElem = vxmlTableElem.ownerDocument.createElement("PERSON")
    Set xmlOperationElem = vxmlTableElem.ownerDocument.createElement("OPERATION")
    xmlOutPersonElem.appendChild xmlOperationElem
    
    Set objIPersonClassDef = New PersonClassdef
    Select Case objXmlAssist.DetermineOperationEx(vxmlTableElem, objIPersonClassDef.LoadPersonData)
        Case booCreate
             Set xmlElement = objIPersonDO.Create(vxmlTableElem)
             xmlOutPersonElem.appendChild xmlElement.getElementsByTagName("PERSONGUID").Item(0)
             xmlOperationElem.Text = "CREATE"
        Case booUpdate
            Call objIPersonDO.Update(vxmlTableElem)
            xmlOperationElem.Text = "UPDATE"
        Case booDelete
            ' Do not call delete method here - PersonGUID has been referred in Alias
            xmlOperationElem.Text = "DELETE"
        Case Else
            ' N/A
    End Select

    Set UpdatePersonForThisAlias = xmlOutPersonElem

    Set objXmlAssist = Nothing
    Set objErrAssist = Nothing
    Set objIPersonDO = Nothing
    Set xmlElement = Nothing
    Set xmlOutPersonElem = Nothing
    Set xmlOperationElem = Nothing

    If Not objContext Is Nothing Then
        objContext.SetComplete
        Set objContext = Nothing
    End If
    Exit Function

UpdatePersonForThisAliasVbErr:

    Set objXmlAssist = Nothing
    Set objIPersonDO = Nothing
    Set xmlElement = Nothing
    Set xmlOutPersonElem = Nothing
    Set xmlOperationElem = Nothing
    
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
        Set objContext = Nothing
    End If

    Set objErrAssist = Nothing

    '  re-raise error for business object to interpret as appropriate
    Err.Raise Err.Number, Err.Source, Err.Description

End Function

Private Function ICustomerDO_GetPersonalDetails(ByVal vXMLRequest As IXMLDOMElement) As IXMLDOMElement
' header ----------------------------------------------------------------------------------
' description:
'   Get the data of Customer - from tables Customer, Customerversion
' pass:
'   vstrXMLRequest  xml Request data stream containing data to which identifies
'   the instance of the persistant data to be retrieved - Customer number and customerversion
' return:
'   GetPersonDetails : string containing XML data stream representation of data retrieved
' Raise Errors: if record not found, raise omiga4RecordNotFound
'------------------------------------------------------------------------------------------
  On Error GoTo ICustomerDO_GetPersonalDetailsVbErr

    Const strFunctionName As String = "ICustomerDO_GetPersonalDetails"

    Dim objContext As ObjectContext
    Set objContext = GetObjectContext()
    
    Dim objXmlAssist As New XMLAssist
    Dim objErrAssist As New ErrAssist
    Dim objSQLAssist As New SQLAssist
    Dim objIClassDef As IomCustClassDef
    
    Dim strSQL As String 'Query to be executed
    Dim strTempCondition As String 'Where condtion to be applied on the table
    Dim strXML As String 'Output (XML string) of the function

    Dim xmlIn As FreeThreadedDOMDocument40
    Dim xmlResponseElem As IXMLDOMElement

    Dim strCustomerNumber As String, intCustomerVersionNumber As Integer

    strCustomerNumber = objXmlAssist.GetTagValue(vXMLRequest, "CUSTOMERNUMBER")
    intCustomerVersionNumber = objXmlAssist.GetTagValue(vXMLRequest, "CUSTOMERVERSIONNUMBER")

    Dim xmlCustomerElem As IXMLDOMElement 'customer element in XMLout - Top most element in XMLOut
    Dim xmlCustVersionElem As IXMLDOMElement 'customer version in XMLOut

    'Add the node for Customer table
    Set objIClassDef = New omCustClassDef
    
    strTempCondition = "CustomerNumber =" & objSQLAssist.FormatString(strCustomerNumber)
    Set xmlCustomerElem = BuildNode("Customer", objIClassDef.LoadCustomerData, strTempCondition)

    'Add the node from customerversion table
    strTempCondition = "CUSTOMERNUMBER =" & objSQLAssist.FormatString(strCustomerNumber) _
                       & " And CUSTOMERVERSIONNUMBER = " & intCustomerVersionNumber
    Set xmlCustVersionElem = BuildNode("CUSTOMERVERSION", _
                            objIClassDef.LoadCustomerVersionData, strTempCondition)
    
    Call AddCustVersionDerivedData(xmlCustVersionElem)

     xmlCustomerElem.appendChild xmlCustVersionElem

    'Assign the Output String
    Set ICustomerDO_GetPersonalDetails = xmlCustomerElem

    'Clear all the objects
    Set objErrAssist = Nothing
    Set objXmlAssist = Nothing

    Set xmlIn = Nothing

    Set xmlCustomerElem = Nothing
    Set xmlCustVersionElem = Nothing

    If Not objContext Is Nothing Then
        objContext.SetComplete
    End If

    Exit Function

ICustomerDO_GetPersonalDetailsVbErr:

    Set objXmlAssist = Nothing
    Set xmlIn = Nothing

    Set xmlCustomerElem = Nothing
    Set xmlCustVersionElem = Nothing

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

    're-raise error for business object to interpret as appropriate
    Err.Raise Err.Number, Err.Source, Err.Description
End Function

Public Function AddCustVersionDerivedData(ByVal vXMLData As IXMLDOMElement) As IXMLDOMElement
' header ----------------------------------------------------------------------------------
' Description :
' This function builds the derived data for Customer Version table and adds to xml node
' ( with customer version data ) passed.
' pass:     Node to append the derived data to
' return:   Resultant node (after appending the derived data)
' Raise Errors:
'------------------------------------------------------------------------------------------
On Error GoTo AddCustVersionDerivedDataVBErr

    Const strFunctionName As String = "AddCustVersionDerivedData"

    Dim objErrAssist As New ErrAssist

    Dim objXmlAssist As New XMLAssist
    Dim objXmlElem As IXMLDOMElement

    ' *****************************************************************************
    '   The following stores Age as derived from date of birth
    ' *****************************************************************************
    Dim strDoB As String
    Dim blnTagExists As Boolean
    strDoB = objXmlAssist.GetTagValue(vXMLData, "DATEOFBIRTH", blnTagExists)
    If blnTagExists Then
        Set objXmlElem = vXMLData.ownerDocument.createElement("AGE")
        If Len(strDoB) > 0 Then
            objXmlElem.Text = GetAge(strDoB)
        End If
        vXMLData.appendChild objXmlElem
    End If

    ' *****************************************************************************
    '   The following stores Initials as derived from FirstForename,
    '   & SecondForename and OtherForeNames
    ' *****************************************************************************

    Dim strInitials As String

    If vXMLData.getElementsByTagName("FIRSTFORENAME").length > 0 Then
        If Len(vXMLData.getElementsByTagName("FIRSTFORENAME").Item(0).Text) > 0 Then
            strInitials = Left(vXMLData.getElementsByTagName("FIRSTFORENAME").Item(0).Text, 1)
        End If
    End If

    If vXMLData.getElementsByTagName("SECONDFORENAME").length > 0 Then
        If Len(vXMLData.getElementsByTagName("SECONDFORENAME").Item(0).Text) > 0 Then
            strInitials = strInitials & Left(vXMLData.getElementsByTagName("SECONDFORENAME").Item(0).Text, 1)
        End If
    End If

    If vXMLData.getElementsByTagName("OTHERFORENAMES").length > 0 Then
        If Len(vXMLData.getElementsByTagName("OTHERFORENAMES").Item(0).Text) > 0 Then
            Dim strOtherForeNames As String
            strOtherForeNames = vXMLData.getElementsByTagName("OTHERFORENAMES").Item(0).Text
            strInitials = strInitials & Left(strOtherForeNames, 1)
            Do While InStr(strOtherForeNames, " ") <> 0
                strOtherForeNames = Right(strOtherForeNames, Len(strOtherForeNames) - InStr(strOtherForeNames, " "))
                strInitials = strInitials & Left(strOtherForeNames, 1)
            Loop
        End If
    End If

    If Len(strInitials) > 0 Then
        Set objXmlElem = vXMLData.ownerDocument.createElement("INITIALS")
        objXmlElem.Text = strInitials
        vXMLData.appendChild objXmlElem
    End If

    Set AddCustVersionDerivedData = vXMLData

    Set objXmlAssist = Nothing
    Set objErrAssist = Nothing
    Set objXmlElem = Nothing

    Exit Function

AddCustVersionDerivedDataVBErr:

'    Set objXmlDoc = Nothing
'    Set objXmlAssist = Nothing
'    Set objXmlElem = Nothing
'
'    AddCustVersionDerivedData = ""
'
'    If objErrAssist.IsSystemError Then
'        objErrAssist.LogError TypeName(Me), strFunctionName, Err.Number, Err.Description
'    End If
'
'    Set objErrAssist = Nothing
'
'    Err.Raise Err.Number, Err.Source, Err.Description
End Function

Private Function BuildNode(ByVal strTable As String, _
                            ByVal xmlClassDef As FreeThreadedDOMDocument40, _
                            ByVal strCondition As String) As IXMLDOMNode
'*Header************************************************************************************
'Description: Builds the node from the table from the records that match the criteria.
'             RAISES ERROR IF NO RECORDS ARE FOUND
'Pass : TableName, Class Definition Template, condtion to be applied on the table during query
'Return: XML String
'Raise Error : oeRecordNotfound
'*******************************************************************************************
On Error GoTo BuildNodeVbErr

    Const strFunctionName As String = "BuildNode"

    Dim objContext As ObjectContext
    Set objContext = GetObjectContext()

    Dim objIADOAssist As IADOAssist
    Dim objIDOAssist As IDOAssist
    Dim objErrAssist As New ErrAssist
    Dim objSQLAssist As New SQLAssist

    Dim strSQL As String, strXML As String
    Dim cmd As New ADODB.Command
    Dim rstThisRecSet As New ADODB.Recordset
    Dim xmlRetNode As IXMLDOMNode

    If objContext Is Nothing Then
        Set objIADOAssist = New ADOAssist
        Set objIDOAssist = New DOAssist
    Else
        Set objIADOAssist = objContext.CreateInstance(gstrBASE_COMPONENT & ".ADOAssist")
        Set objIDOAssist = objContext.CreateInstance(gstrBASE_COMPONENT & ".DOAssist")
    End If

    strSQL = "SELECT * FROM " & strTable
    If strCondition <> "" Then
        strSQL = strSQL & " WHERE " & strCondition
    End If

    cmd.ActiveConnection = objIADOAssist.GetConnStr
    cmd.CommandText = strSQL
    Set rstThisRecSet = cmd.Execute(, , adCmdText)

    If rstThisRecSet.EOF Then 'if no matching record exists return empty string
        'Raise Error
        objErrAssist.ThrowError TypeName(Me), strFunctionName, oeRecordNotFound, _
                                 strTable

    Else
        'Add the details from the recordset to XML string
        Set xmlRetNode = objIDOAssist.GetXMLFromRecordSet(rstThisRecSet, xmlClassDef)
    End If
        rstThisRecSet.Close
    
    Set BuildNode = xmlRetNode.cloneNode(True)

    Set objIADOAssist = Nothing
    Set objIDOAssist = Nothing
    Set objErrAssist = Nothing
    Set rstThisRecSet = Nothing
    Set cmd = Nothing
    
    Set xmlRetNode = Nothing
    
    Exit Function
BuildNodeVbErr:

    ' Close the recordset if still open
        If Not rstThisRecSet Is Nothing Then
                If rstThisRecSet.State = adStateOpen Then
                        rstThisRecSet.Close
                End If
        End If

    Set objIADOAssist = Nothing
    Set objIDOAssist = Nothing
    Set rstThisRecSet = Nothing
    Set cmd = Nothing
    Set xmlRetNode = Nothing
    
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
        Set objContext = Nothing
    End If

    Set objErrAssist = Nothing

    're-raise error for business object to interpret as appropriate
    Err.Raise Err.Number, Err.Source, Err.Description
End Function

Private Function ICustomerDO_GetAlias(ByVal vXMLRequest As IXMLDOMElement) As IXMLDOMElement
' header ----------------------------------------------------------------------------------
' description:
'   Get the data of Alias - from tables Alias, Person
' pass:
'   vXMLRequest     xml Request node which identifies the instance of the persistant
'                   data to be retrieved (Alias)
' return:
' Raise Errors: if record not found, raise omiga4RecordNotFound
'------------------------------------------------------------------------------------------
On Error GoTo ICustomerDO_GetAliasVbErr

    Const strFunctionName As String = "ICustomerDO_GetAlias"
    
    Dim objContext As ObjectContext
    Set objContext = GetObjectContext()
    
    Dim objXmlAssist As New XMLAssist
    Dim objErrAssist As New ErrAssist
    Dim objSQLAssist As New SQLAssist
    
    Dim objIClassDef As IomCustClassDef
    Dim objIPersonClassDef As IPersonClassDef
   
    Dim strSQL As String 'Query to be executed
    Dim xmlAliasDef As FreeThreadedDOMDocument40, xmlPersonDef As FreeThreadedDOMDocument40
    Dim xmlResponse As IXMLDOMNode

    Dim strCustomerNumber As String
    Dim strCustomerVersionNumber As String
    Dim strAliasSequenceNumber As String
    Dim strPersonGuid As String

    'Fetch the values of all the columns in the Primary Key and PersonGuid.
    strCustomerNumber = objXmlAssist.GetMandatoryElementText(vXMLRequest, "CUSTOMERNUMBER")
    strCustomerVersionNumber = objXmlAssist.GetMandatoryElementText(vXMLRequest, _
                                                          "CUSTOMERVERSIONNUMBER")
    strAliasSequenceNumber = objXmlAssist.GetMandatoryElementText(vXMLRequest, _
                                                        "ALIASSEQUENCENUMBER")

    Dim objIDOAssist As DOAssist
    Dim objIADOAssist As ADOAssist

    If objContext Is Nothing Then
        Set objIDOAssist = New DOAssist
        Set objIADOAssist = New ADOAssist
    Else
        Set objIDOAssist = objContext.CreateInstance(gstrBASE_COMPONENT & ".DOAssist")
        Set objIADOAssist = objContext.CreateInstance(gstrBASE_COMPONENT & ".ADOAssist")
    End If

    'Build the SQL string to be executed
    strSQL = "SELECT * FROM FINDALIASLIST "

    strSQL = strSQL & " WHERE CUSTOMERNUMBER = " & objSQLAssist.FormatString(strCustomerNumber) _
                    & " And CUSTOMERVERSIONNUMBER = " & strCustomerVersionNumber _
                    & " And ALIASSEQUENCENUMBER = " & strAliasSequenceNumber

    'Open the recordset with the SQL string built above
    Dim cmd As New ADODB.Command
    cmd.ActiveConnection = objIADOAssist.GetConnStr
    cmd.CommandText = strSQL

    Dim rstThisRecSet As New ADODB.Recordset
    Set rstThisRecSet = cmd.Execute(, , adCmdText)

    If rstThisRecSet.EOF And rstThisRecSet.BOF Then
        ' raise application error to be interpreted by calling object
        objErrAssist.ThrowError TypeName(Me), strFunctionName, oeRecordNotFound
    End If

    ' Define classes for the Templates
    Set objIPersonClassDef = New PersonClassdef
    Set objIClassDef = New omCustClassDef
        
    Set xmlResponse = vXMLRequest.ownerDocument.createElement("RESPONSE")

    'Add Alias XML
    Call objIDOAssist.GetXMLFromRecordsetEx(rstThisRecSet, objIClassDef.LoadAliasData, _
                                               xmlResponse)

    'Add Person XML string, If the PersonGuid
    If Not IsNull(rstThisRecSet.Fields("PersonGuid")) Then
        Call objIDOAssist.GetXMLFromRecordsetEx(rstThisRecSet, objIPersonClassDef.LoadPersonData, _
                                                   xmlResponse.firstChild)
    End If
        rstThisRecSet.Close
    Set ICustomerDO_GetAlias = xmlResponse.selectSingleNode("ALIAS")
    
    'Clear all the objects
    Set rstThisRecSet = Nothing
    Set cmd = Nothing

    Set objIDOAssist = Nothing
    Set objIADOAssist = Nothing
    Set objIPersonClassDef = Nothing
    Set objIClassDef = Nothing

    Set objErrAssist = Nothing
    Set objXmlAssist = Nothing
    Set xmlResponse = Nothing

    If Not objContext Is Nothing Then
        objContext.SetComplete
        Set objContext = Nothing
    End If
    Set objContext = Nothing

    Exit Function

ICustomerDO_GetAliasVbErr:

    ' Close the recordset if still open
        If Not rstThisRecSet Is Nothing Then
                If rstThisRecSet.State = adStateOpen Then
                        rstThisRecSet.Close
                End If
        End If

    Set objXmlAssist = Nothing
    Set objIDOAssist = Nothing
    Set objIADOAssist = Nothing
    
    Set rstThisRecSet = Nothing
    Set cmd = Nothing
    
    Set objIPersonClassDef = Nothing
    Set objIClassDef = Nothing
    Set xmlResponse = Nothing

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
        Set objContext = Nothing
    End If

    Set objErrAssist = Nothing

    're-raise error for business object to interpret as appropriate
    Err.Raise Err.Number, Err.Source, Err.Description

End Function

Public Function ICustomerDO_FindAliasList(ByVal vXMLRequestElem As IXMLDOMElement) As IXMLDOMElement
' header ----------------------------------------------------------------------------------
' description:
'   Get the data the Alias and Person that match the criterion as passed to the function
' pass:
'   vxmlRequestElem  xml Request node which identifies the creiterion to fetch the data from database
' return: xml node containing the data retreived from the database
' Raise Errors: if record not found, raise omiga4RecordNotFound
'------------------------------------------------------------------------------------------
On Error GoTo ICustomerDO_FindAliasListVbErr

    Const strFunctionName As String = "ICustomerDO_FindAliasList"

    Dim objContext As ObjectContext
    Set objContext = GetObjectContext()

    Dim objXmlAssist As New XMLAssist
    Dim objErrAssist As New ErrAssist
    Dim objSQLAssist As New SQLAssist
    
    Dim objIADOAssist As IADOAssist
    Dim objIDOAssist As IDOAssist
    Dim objIClassDef As IomCustClassDef
    Dim objIPersonClassDef As IPersonClassDef
    Dim xmlAliasDef As FreeThreadedDOMDocument40, xmlPersonDef As FreeThreadedDOMDocument40
    
    Dim xmlAliasPersonListElem As IXMLDOMElement
    Dim xmlAliasPersonElem As IXMLDOMElement
    Dim xmlAliasNode As IXMLDOMNode
    Dim xmlResponseNode As IXMLDOMNode
    Dim xmlNodeList As IXMLDOMNodeList, xmlNode As IXMLDOMNode

    Dim strRecCondition As String, strFinalCondition As String, strSQL As String
    Dim colValueids As Collection
    Dim strValueid As String

    'Retrieve the values of the input params passed thru request string
    Dim strCustomerNumber As String
    Dim strCustomerVersionNumber As String
    Dim strAliasSequenceNumber As String
    Dim blnAliasRecordsOnly As Boolean
    Dim blnDeclaredRecordsOnly As Boolean

    If objContext Is Nothing Then
        Set objIADOAssist = New ADOAssist
        Set objIDOAssist = New DOAssist
    Else
        Set objIADOAssist = objContext.CreateInstance(gstrBASE_COMPONENT & ".ADOAssist")
        Set objIDOAssist = objContext.CreateInstance(gstrBASE_COMPONENT & ".DOAssist")
    End If
       ' BMIDS693 DRC - CreditSearch START
    blnAliasRecordsOnly = (objXmlAssist.GetElementText(vXMLRequestElem, "ALIASRECORDSONLY") = "1")
    blnDeclaredRecordsOnly = (objXmlAssist.GetElementText(vXMLRequestElem, "DECLAREDRECORDSONLY") = "1")
    ' BMIDS693 DRC - CreditSearch END

    ' Loop through the request string and build the condtion to applied on the query
    ' to the database
    Set xmlNodeList = vXMLRequestElem.getElementsByTagName("ALIAS")
    For Each xmlNode In xmlNodeList

        strCustomerNumber = objXmlAssist.GetTagValue(xmlNode, "CUSTOMERNUMBER")
        strCustomerVersionNumber = objXmlAssist.GetTagValue(xmlNode, "CUSTOMERVERSIONNUMBER")
        strAliasSequenceNumber = objXmlAssist.GetTagValue(xmlNode, "ALIASSEQUENCENUMBER")
               strRecCondition = ""
        If strCustomerNumber <> "" Then
            strRecCondition = strRecCondition _
                        & " CUSTOMERNUMBER = " & objSQLAssist.FormatString(strCustomerNumber)
        End If

        If strCustomerVersionNumber <> "" Then
            If strRecCondition <> "" Then strRecCondition = strRecCondition & " And "

            strRecCondition = strRecCondition _
                            & " CUSTOMERVERSIONNUMBER = " & strCustomerVersionNumber
        End If

        If strAliasSequenceNumber <> "" Then
            If strRecCondition <> "" Then strRecCondition = strRecCondition & " AND "

            strRecCondition = strRecCondition _
                            & " ALIASSEQUENCENUMBER = " & strAliasSequenceNumber
        End If
        
        If blnAliasRecordsOnly And blnDeclaredRecordsOnly Then
           Set colValueids = New Collection
           comboAssistEx.GetValueIdsForValidationType "AliasType", "AL", colValueids
           strValueid = CStr(colValueids.Item(1))
           If strRecCondition <> "" Then strRecCondition = strRecCondition & " AND "
            strRecCondition = strRecCondition _
                            & " (CREDITSEARCH is null OR " _
                            & " CREDITSEARCH <> 1) AND ALIASTYPE = " & strValueid
        End If

        If strFinalCondition <> "" Then strFinalCondition = strFinalCondition & " OR "

        If strRecCondition <> "" Then
            strFinalCondition = strFinalCondition & "(" & strRecCondition & ")"
        End If

    Next xmlNode

    'Build the SQL string to be executed
    strSQL = "SELECT * FROM FINDALIASLIST "

    If strFinalCondition <> "" Then
        strSQL = strSQL & " Where " & strFinalCondition
    End If

    'Open the recordset with the SQL string built above
    Dim cmd As New ADODB.Command
    cmd.ActiveConnection = objIADOAssist.GetConnStr
    cmd.CommandText = strSQL

    Dim rstThisRecSet As New ADODB.Recordset
    Set rstThisRecSet = cmd.Execute(, , adCmdText)

    If rstThisRecSet.EOF And rstThisRecSet.BOF Then
        ' raise application error to be interpreted by calling object
        objErrAssist.ThrowError TypeName(Me), strFunctionName, oeRecordNotFound
    End If

    'Define the Objects required for Templates and build templates
    Set objIClassDef = New omCustClassDef
    Set objIPersonClassDef = New PersonClassdef
    
    Set xmlAliasDef = objIClassDef.LoadAliasData
    Set xmlPersonDef = objIPersonClassDef.LoadPersonData
    
    'Add the outer most element to Response Node
    Set xmlResponseNode = vXMLRequestElem.ownerDocument.createElement("RESPONSE")
    Set xmlAliasPersonListElem = xmlResponseNode.ownerDocument.createElement("ALIASPERSONLIST")
    xmlResponseNode.appendChild xmlAliasPersonListElem

    'Build the output XML node
    'AD 19/10/01 Removed the MoveFirst as it's a ForwardOnly cursor. SYS2718.

    'Navigate through the record set and add the respective node to Response
    Do While Not rstThisRecSet.EOF
        Set xmlAliasPersonElem = xmlAliasPersonListElem.ownerDocument.createElement("ALIASPERSON")
        xmlAliasPersonListElem.appendChild xmlAliasPersonElem

        'Add Alias XML node
        Set xmlAliasNode = _
            objIDOAssist.GetXMLFromRecordSet(rstThisRecSet, xmlAliasDef, xmlAliasPersonElem)

        'Add Person XML string, If the PersonGuid is not null
        If Not IsNull(rstThisRecSet.Fields("PersonGuid")) Then
            Call objIDOAssist.GetXMLFromRecordSet(rstThisRecSet, xmlPersonDef, xmlAliasNode)
        End If
        rstThisRecSet.MoveNext
    Loop
    rstThisRecSet.Close

    Set ICustomerDO_FindAliasList = xmlResponseNode.selectSingleNode("ALIASPERSONLIST")
    'Clear all the objects
    Set rstThisRecSet = Nothing
    Set cmd = Nothing

    Set objIDOAssist = Nothing
    Set objIADOAssist = Nothing
    Set objErrAssist = Nothing
    Set objXmlAssist = Nothing
    Set objSQLAssist = Nothing
    
    Set objIClassDef = Nothing
    Set objIPersonClassDef = Nothing
    Set xmlAliasDef = Nothing
    Set xmlPersonDef = Nothing

    Set xmlAliasPersonListElem = Nothing
    Set xmlAliasPersonElem = Nothing
    Set xmlAliasNode = Nothing
    Set xmlResponseNode = Nothing
    
    If Not objContext Is Nothing Then
        objContext.SetComplete
        Set objContext = Nothing
    End If

    Exit Function

ICustomerDO_FindAliasListVbErr:

    ' Close the recordset if still open
        If Not rstThisRecSet Is Nothing Then
                If rstThisRecSet.State = adStateOpen Then
                        rstThisRecSet.Close
                End If
        End If

    Set rstThisRecSet = Nothing
    Set cmd = Nothing
    
    Set objIADOAssist = Nothing
    Set objIDOAssist = Nothing
    Set objXmlAssist = Nothing
    Set objSQLAssist = Nothing
        
    Set objIClassDef = Nothing
    Set objIPersonClassDef = Nothing
    Set xmlAliasDef = Nothing
    Set xmlPersonDef = Nothing

    Set xmlAliasPersonListElem = Nothing
    Set xmlAliasPersonElem = Nothing
    Set xmlAliasNode = Nothing
    Set xmlResponseNode = Nothing

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
        Set objContext = Nothing
    End If

    Set objErrAssist = Nothing

    '   re-raise error for business object to interpret as appropriate
    Err.Raise Err.Number, Err.Source, Err.Description

End Function

Private Function ICustomerDO_FindCustomerAddressList(ByVal vXMLRequest As IXMLDOMElement, _
                    Optional ByVal blnOrderBySequenceNUmber As Boolean = False) As IXMLDOMElement
' header ----------------------------------------------------------------------------------
' description:
'   Get the data from tables CustomerAddress and Address that match the criterion
' pass:
'   vstrXMLRequest  xml Request data stream containing data which identifies
'                   the creiterion to be used in the Query
' return:
'   FindCustomerAddressList  string containing XML data stream representation
'    of data retrieved in the format
'   <CUSTOMERADDRESSLIST>
'       <CUSTOMERADDRESS>
'           <CUSTOMERNUMBER> ... </CUSTOMERNUMBER>
'           <CUSTOMERVERSIONNUMBER> .... </CUSTOMERVERSIONNUMBER>
'           ........
'          <ADDRESS>
'               <ADDRESSGUID> ... </ADDRESSGUID>
'               .......
'         </ADDRESS>
'       </CUSTOMERADDRESS>
'  </CUSTOMERADDRESSLIST>
' Raise Errors: if record not found, raise omiga4RecordNotFound
'------------------------------------------------------------------------------------------
On Error GoTo ICustomerDO_FindCustomerAddressListVbErr

    Const strFunctionName As String = "ICustomerDO_FindCustomerAddressList"

    Dim objXmlAssist As New XMLAssist
    Dim objErrAssist As New ErrAssist
    Dim objSQLAssist As New SQLAssist

    Dim objContext As ObjectContext
    Set objContext = GetObjectContext()

    Dim strSQL As String
    Dim strFinalCondition As String ' Condition to be added to select string before fetch data
    Dim strRecCondition As String ' Condition based on single row in xmlRequest
    
    Dim objIClassDef As IomCustClassDef, objIAddressClassDef As IAddressClassDef
    Dim xmlAddressDef As FreeThreadedDOMDocument40, xmlCustomerAddressDef As FreeThreadedDOMDocument40
    
    Dim xmlCustomerAddressListElem As IXMLDOMElement
    Dim xmlNodeList As IXMLDOMNodeList
    Dim xmlNode As IXMLDOMNode

    Dim objIADOAssist As IADOAssist
    Dim objIDOAssist As IDOAssist
    Dim objIAddressDO As New IAddressDO

    If objContext Is Nothing Then
        Set objIADOAssist = New ADOAssist
        Set objIDOAssist = New DOAssist
    Else
        Set objIADOAssist = objContext.CreateInstance(gstrBASE_COMPONENT & ".ADOAssist")
        Set objIDOAssist = objContext.CreateInstance(gstrBASE_COMPONENT & ".DOAssist")
    End If

    ' Build the criteria to be applied on the view to fetch data
    
    If vXMLRequest.nodeName = "CUSTOMERADDRESS" Then
        Set xmlNode = vXMLRequest
        strFinalCondition = BuildConditionForFindCustomerAddressList(xmlNode)
    Else
        Set xmlNodeList = vXMLRequest.getElementsByTagName("CUSTOMERADDRESS")
        
        For Each xmlNode In xmlNodeList
            strRecCondition = BuildConditionForFindCustomerAddressList(xmlNode)

            If strRecCondition <> "" Then
                If strFinalCondition <> "" Then strFinalCondition = strFinalCondition & " OR "
                
                strFinalCondition = strFinalCondition & "(" & strRecCondition & ")"
            End If
        Next xmlNode
    End If
    
    strSQL = "SELECT * FROM FINDCUSTOMERADDRESSLIST "
    If strFinalCondition <> "" Then
        strSQL = strSQL & " Where " & strFinalCondition
    End If
    
    ' SR - AQR/SYS0390 - sort the data on addressType, DateMovedIn
    ' SR 02/06/00 SYS0805 - Use optional parameter to decide the order by clause to be used
    If blnOrderBySequenceNUmber Then
        strSQL = strSQL & " Order By CUSTOMERADDRESSSEQUENCENUMBER ASC "
    Else
        'BM0029 MDC 17/12/2002 - Sort by Customer Order too
        'NB: This change will work because BMids create a new customer version for each app.
        'If in the future customer versions can be linked to multiple apps, the app number
        'would need to be passed in also.
        'strSQL = strSQL & " Order By ADDRESSTYPE ASC, DATEMOVEDIN Desc "
        strSQL = strSQL & " Order By ADDRESSTYPE ASC, CUSTOMERORDER ASC, DATEMOVEDIN Desc "
        'BM0029 MDC 17/12/2002 - End
    End If
    
    'Open the recordset with the SQL string built above
    Dim cmd As New ADODB.Command
    cmd.ActiveConnection = objIADOAssist.GetConnStr
    cmd.CommandText = strSQL

    Dim rstThisRecSet As New ADODB.Recordset
    Set rstThisRecSet = cmd.Execute(, , adCmdText)

    If rstThisRecSet.EOF And rstThisRecSet.BOF Then
        ' raise application error to be interpreted by calling object
        objErrAssist.ThrowError TypeName(Me), strFunctionName, oeRecordNotFound
    End If

    ' DEfine the classes required for class definitions
    Set objIClassDef = New omCustClassDef
    Set objIAddressClassDef = New AddressClassDef
    
    Set xmlAddressDef = objIAddressClassDef.LoadAddressData
    Set xmlCustomerAddressDef = objIClassDef.LoadCustomerAddressData

     Set xmlCustomerAddressListElem = vXMLRequest.ownerDocument.createElement("CUSTOMERADDRESSLIST")
    'Navigate through the record set and add the xml node to list element
    'AD 19/10/01 Removed the MoveFirst as it's a ForwardOnly cursor. SYS2718.
    Do While Not rstThisRecSet.EOF
        Set xmlNode = objIDOAssist.GetXMLFromRecordSet(rstThisRecSet, xmlCustomerAddressDef, _
                                               xmlCustomerAddressListElem)
        'Add Address XML string, If the AddressGuid is not null
        If Not IsNull(rstThisRecSet.Fields("AddressGuid")) Then
            Call objIDOAssist.GetXMLFromRecordSet(rstThisRecSet, xmlAddressDef, xmlNode)
        End If
        rstThisRecSet.MoveNext
    Loop
        rstThisRecSet.Close
    Set ICustomerDO_FindCustomerAddressList = xmlCustomerAddressListElem
    
    'Clear all the objects
    Set rstThisRecSet = Nothing
    Set cmd = Nothing

    Set objIDOAssist = Nothing
    Set objIADOAssist = Nothing
    Set objIAddressDO = Nothing

    Set objErrAssist = Nothing
    Set objXmlAssist = Nothing
    Set objSQLAssist = Nothing

    Set xmlCustomerAddressListElem = Nothing
    Set xmlNode = Nothing
    Set xmlNodeList = Nothing
    Set objIClassDef = Nothing
    Set objIAddressClassDef = Nothing
    Set xmlAddressDef = Nothing
    Set xmlCustomerAddressDef = Nothing

    If Not objContext Is Nothing Then
        objContext.SetComplete
        Set objContext = Nothing
    End If

    Exit Function

ICustomerDO_FindCustomerAddressListVbErr:
    
    ' Close the recordset if still open
        If Not rstThisRecSet Is Nothing Then
                If rstThisRecSet.State = adStateOpen Then
                        rstThisRecSet.Close
                End If
        End If

    Set rstThisRecSet = Nothing
    Set cmd = Nothing
        Set objIDOAssist = Nothing
    Set objIADOAssist = Nothing
    Set objIAddressDO = Nothing

    Set objXmlAssist = Nothing
    Set objSQLAssist = Nothing

    Set xmlCustomerAddressListElem = Nothing
    Set xmlNode = Nothing
    Set xmlNodeList = Nothing
    Set objIClassDef = Nothing
    Set objIAddressClassDef = Nothing
    Set xmlAddressDef = Nothing
    Set xmlCustomerAddressDef = Nothing
    
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
        Set objContext = Nothing
    End If

    Set objErrAssist = Nothing

    '   re-raise error for business object to interpret as appropriate
    Err.Raise Err.Number, Err.Source, Err.Description

End Function

Private Function BuildConditionForFindCustomerAddressList(ByVal xmlNode As IXMLDOMNode) As String
' header ----------------------------------------------------------------------------------
' description:
'   Builds the condition to be used in SELECT statement for FindCustomerAddressList
' pass:
'   xmlNode
' Raise Errors:
'------------------------------------------------------------------------------------------
On Error GoTo errhandler
    Const strFunctionName As String = "BuildConditionForFindCustomerAddressList"

    Dim objXmlAssist As New XMLAssist
    Dim objErrAssist As New ErrAssist
    Dim objSQLAssist As New SQLAssist
    
    Dim strRecCondition As String
    Dim strCustomerNumber As String, strCustomerVersionNumber As String
    Dim strAddressType As String
    
    strCustomerNumber = objXmlAssist.GetTagValue(xmlNode, "CUSTOMERNUMBER")
    strCustomerVersionNumber = objXmlAssist.GetTagValue(xmlNode, "CUSTOMERVERSIONNUMBER")
    strAddressType = objXmlAssist.GetTagValue(xmlNode, "ADDRESSTYPE")

    If strCustomerNumber <> "" Then
        strRecCondition = " CUSTOMERNUMBER = " & objSQLAssist.FormatString(strCustomerNumber)
    End If

    If strCustomerVersionNumber <> "" Then
        If strRecCondition <> "" Then strRecCondition = strRecCondition & " And "
        strRecCondition = strRecCondition _
                        & " CUSTOMERVERSIONNUMBER = " & strCustomerVersionNumber
    End If

    If strAddressType <> "" Then
        If strRecCondition <> "" Then strRecCondition = strRecCondition & " And "
        strRecCondition = strRecCondition & " ADDRESSTYPE = " & strAddressType
    End If
    
    BuildConditionForFindCustomerAddressList = strRecCondition
    
    Set objXmlAssist = Nothing
    Set objErrAssist = Nothing
    Set objSQLAssist = Nothing
    
    Exit Function
    
errhandler:
        
    Set objXmlAssist = Nothing
    Set objSQLAssist = Nothing
    objErrAssist.AddToErrSource (strFunctionName)

    If objErrAssist.IsSystemError = True Then
        objErrAssist.LogError TypeName(Me), strFunctionName, Err.Number, Err.Description
    End If

    Set objErrAssist = Nothing

    '   re-raise error for business object to interpret as appropriate
    Err.Raise Err.Number, Err.Source, Err.Description
End Function

Private Function ICustomerDO_GetCustomerAddress(ByVal vXMLRequest As IXMLDOMElement) As IXMLDOMElement
' header ----------------------------------------------------------------------------------
' description:
'   Get the data of Customer Address (from tables Address, CurrentProperty, Home Insurance)
' pass:
'   vXMLRequest  xml Request node which identifies the instance of the persistant data
'                to be retrieved
' return:
'   Response Node
'   The element added to Response node is of the the format
'   <CUSTOMERADDRESS>
'       <CUSTOMERNUMBER> ... </CUSTOMERNUMBER>
'       ........
'       <ADDRESS>
'           <ADDRESSGUID>.....</ADDRESSGUID>
'           ....
'        </ADDRESS>
'       <CURRENTPROPERTY>
'           .......
'       </CURRENTPROPERTY>
'       <HOMEINSURANCE>
'           .......
'       <HOMEINSURANCE>
'   </CUSTOMERADDRESS>
' Raise Errors: if record not found, raise omiga4RecordNotFound
'------------------------------------------------------------------------------------------
  On Error GoTo GetCustomerAddressVbErr

    Const strFunctionName As String = "ICustomerDO_GetCustomerAddress"

    Dim objXmlAssist As New XMLAssist
    Dim objErrAssist As New ErrAssist
    Dim objSQLAssist As New SQLAssist

    Dim objContext As ObjectContext
    Set objContext = GetObjectContext()

    Dim strSQL As String 'Query to be executed
    Dim strCondition As String

    Dim xmlNode As IXMLDOMNode
    Dim xmlCustomerAddressElem As IXMLDOMElement
    Dim xmlDefTemp As FreeThreadedDOMDocument40 'Class Definition of Table
    Dim objIClassDef As IomCustClassDef
    Dim objIAddressClassDef As IAddressClassDef

    Dim strCustomerNumber As String, strCustomerVersionNumber As String
    Dim strCustAddrSequenceNumber As String

    'Fetch the values of all the columns in the Primary Key and PersonGuid.
    strCustomerNumber = vXMLRequest.selectSingleNode("CUSTOMERNUMBER").Text
    strCustomerVersionNumber = vXMLRequest.selectSingleNode("CUSTOMERVERSIONNUMBER").Text
    strCustAddrSequenceNumber = _
                    vXMLRequest.selectSingleNode("CUSTOMERADDRESSSEQUENCENUMBER").Text

    Dim objIDOAssist As IDOAssist
    Dim objIADOAssist As IADOAssist

    If objContext Is Nothing Then
        Set objIDOAssist = New DOAssist
        Set objIADOAssist = New ADOAssist
    Else
        Set objIDOAssist = objContext.CreateInstance(gstrBASE_COMPONENT & ".DOAssist")
        Set objIADOAssist = objContext.CreateInstance(gstrBASE_COMPONENT & ".ADOAssist")
    End If

    'Build the SQL string to be executed
    strSQL = "SELECT * FROM GETCUSTOMERADDRESS "

    strCondition = " CUSTOMERNUMBER = " & objSQLAssist.FormatString(strCustomerNumber) _
                    & " AND CUSTOMERVERSIONNUMBER = " & strCustomerVersionNumber _
                    & " AND CUSTOMERADDRESSSEQUENCENUMBER = " & strCustAddrSequenceNumber

    strSQL = strSQL & " WHERE " & strCondition

    'Open the recordset with the SQL string built above
    Dim cmd As New ADODB.Command
    cmd.ActiveConnection = objIADOAssist.GetConnStr
    cmd.CommandText = strSQL

    Dim rstThisRecSet As New ADODB.Recordset
    Set rstThisRecSet = cmd.Execute(, , adCmdText)

    If rstThisRecSet.EOF And rstThisRecSet.BOF Then
        ' raise application error to be interpreted by calling object
        objErrAssist.ThrowError TypeName(Me), strFunctionName, oeRecordNotFound
    End If

    'Build the class definition combining all the required tables in hierarchy
    Set objIClassDef = New omCustClassDef
    Set objIAddressClassDef = New AddressClassDef
    
    Dim strXML As String
    Set xmlDefTemp = objIClassDef.LoadCustomerAddressData
    strXML = "<TABLENAME>GETCUSTOMERADDRESS"
    strXML = strXML & xmlDefTemp.xml

    If Not IsNull(rstThisRecSet.Fields("AddressGuid")) Then
        Set xmlDefTemp = objIAddressClassDef.LoadAddressData
        strXML = strXML & xmlDefTemp.xml
    End If

    Set xmlDefTemp = objIClassDef.LoadCurrentPropertyData
    strXML = strXML & xmlDefTemp.xml

    Set xmlDefTemp = objIClassDef.LoadHomeInsuranceData
    strXML = strXML & xmlDefTemp.xml
    strXML = strXML & "</TABLENAME>"

    Set xmlDefTemp = objXmlAssist.load(strXML, TypeName(Me), strFunctionName)

    ' Add the XML string (built based on the hierarchial classdefinition and recordset)
    ' to the response
    Set xmlCustomerAddressElem = objIDOAssist.GetXMLFromRecordSet(rstThisRecSet, xmlDefTemp)
    rstThisRecSet.Close

    Set xmlNode = xmlCustomerAddressElem.selectSingleNode("CUSTOMERADDRESS")
    objXmlAssist.PromoteNode xmlCustomerAddressElem, xmlNode
    Set ICustomerDO_GetCustomerAddress = xmlCustomerAddressElem.cloneNode(True)
    
    'Clear all the objects
    Set rstThisRecSet = Nothing
    Set cmd = Nothing

    Set objIDOAssist = Nothing
    Set objIADOAssist = Nothing
    Set objIClassDef = Nothing
    Set objIAddressClassDef = Nothing
    Set xmlDefTemp = Nothing
       
    Set objErrAssist = Nothing
    Set objXmlAssist = Nothing

    Set xmlCustomerAddressElem = Nothing

    If Not objContext Is Nothing Then
        objContext.SetComplete
        Set objContext = Nothing
    End If

    Exit Function

GetCustomerAddressVbErr:

    ' Close the recordset if still open
        If Not rstThisRecSet Is Nothing Then
                If rstThisRecSet.State = adStateOpen Then
                        rstThisRecSet.Close
                End If
        End If

    Set objXmlAssist = Nothing
    Set xmlCustomerAddressElem = Nothing
    Set xmlDefTemp = Nothing

    Set objIDOAssist = Nothing
    Set objIADOAssist = Nothing
    Set rstThisRecSet = Nothing
    Set cmd = Nothing
    Set objIClassDef = Nothing
    Set objIAddressClassDef = Nothing

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
        Set objContext = Nothing
    End If

    Set objErrAssist = Nothing

    're-raise error for business object to interpret as appropriate
    Err.Raise Err.Number, Err.Source, Err.Description

End Function

Private Sub ICustomerDO_UpdateContactHistory(ByVal vxmlTableElement As IXMLDOMElement)
' header ----------------------------------------------------------------------------------
' description:
'   Update a single instance of the persistant data associated with this
'   data object
' pass:
'   vxmlTableElement  xml element containing the request
' return:                n/a
'------------------------------------------------------------------------------------------
On Error GoTo UpdateContactHistoryVbErr

    Const strFunctionName As String = "ICustomerDO_UpdateContactHistory"
    
         
    Dim xmlClassDefDoc As FreeThreadedDOMDocument40
    Dim objIClassDef As IomCustClassDef
       
    If m_objContext Is Nothing Then
        m_objErrAssist.ThrowError TypeName(Me), strFunctionName, oeMTSNotFound
    End If
       
    Set objIClassDef = New omCustClassDef
    Set xmlClassDefDoc = objIClassDef.LoadContactHistoryData()
    m_objIDOAssist.Update vxmlTableElement, xmlClassDefDoc
    
    Set objIClassDef = Nothing
    Set xmlClassDefDoc = Nothing
    
    m_objContext.SetComplete
        
    
    Exit Sub

UpdateContactHistoryVbErr:
            
    
    Set objIClassDef = Nothing
    Set xmlClassDefDoc = Nothing
    
    m_objErrAssist.AddToErrSource (strFunctionName)
    
    If m_objErrAssist.IsSystemError = True Then
        m_objErrAssist.LogError TypeName(Me), strFunctionName, Err.Number, Err.Description
    End If
    
    If m_objErrAssist.IsSystemError Then
        m_objContext.SetAbort
    Else
        m_objContext.SetComplete
    End If

    '   re-raise error for business object to interpret as appropriate
    Err.Raise Err.Number, Err.Source, Err.Description

End Sub

Private Sub ICustomerDO_UpdateCurrentProperty(ByVal vxmlTableElement As IXMLDOMElement)
' header ----------------------------------------------------------------------------------
' description:
'   Update a single instance of the persistant data associated with this
'   data object on the CurrentProperty Table
' pass:
'   vxmlTableElement  xml element containing the request
' return:                n/a
'------------------------------------------------------------------------------------------
On Error GoTo UpdateCurrentPropertyVbErr

    Const strFunctionName As String = "ICustomerDO_UpdateCurrentProperty"
            
    Dim xmlClassDefDoc As FreeThreadedDOMDocument40
    Dim objIClassDef As IomCustClassDef
    
    
    Set objIClassDef = New omCustClassDef
    Set xmlClassDefDoc = objIClassDef.LoadCurrentPropertyData()
    m_objIDOAssist.Update vxmlTableElement, xmlClassDefDoc
    
    Set objIClassDef = Nothing
    Set xmlClassDefDoc = Nothing
    
    If Not m_objContext Is Nothing Then
        m_objContext.SetComplete
    End If
    
    Exit Sub

UpdateCurrentPropertyVbErr:
        
    Set objIClassDef = Nothing
    Set xmlClassDefDoc = Nothing
    
    m_objErrAssist.AddToErrSource (strFunctionName)
    
    If m_objErrAssist.IsSystemError = True Then
        m_objErrAssist.LogError TypeName(Me), strFunctionName, Err.Number, Err.Description
    End If
    
    If Not m_objContext Is Nothing Then
        If m_objErrAssist.IsSystemError Then
            m_objContext.SetAbort
        Else
            m_objContext.SetComplete
        End If
    End If

    '   re-raise error for business object to interpret as appropriate
    Err.Raise Err.Number, Err.Source, Err.Description

End Sub

Private Sub ICustomerDO_UpdateCustomer(ByVal vxmlTableElement As IXMLDOMElement)

On Error GoTo UpdateCustomerVbErr

    Const strFunctionName As String = "ICustomerDO_UpdateCustomer"
            
    Dim xmlClassDefDoc As FreeThreadedDOMDocument40
    Dim objIClassDef As IomCustClassDef
    
    Set objIClassDef = New omCustClassDef
    Set xmlClassDefDoc = objIClassDef.LoadCustomerData()
    m_objIDOAssist.Update vxmlTableElement, xmlClassDefDoc
    
    Set objIClassDef = Nothing
    Set xmlClassDefDoc = Nothing
    
    If Not m_objContext Is Nothing Then
        m_objContext.SetComplete
    End If
    
    Exit Sub

UpdateCustomerVbErr:
        
    Set objIClassDef = Nothing
    Set xmlClassDefDoc = Nothing
    
    m_objErrAssist.AddToErrSource (strFunctionName)
    
    If m_objErrAssist.IsSystemError = True Then
        m_objErrAssist.LogError TypeName(Me), strFunctionName, Err.Number, Err.Description
    End If
    
    If Not m_objContext Is Nothing Then
        If m_objErrAssist.IsSystemError Then
            m_objContext.SetAbort
        Else
            m_objContext.SetComplete
        End If
    End If

    '   re-raise error for business object to interpret as appropriate
    Err.Raise Err.Number, Err.Source, Err.Description

End Sub

Private Sub ICustomerDO_UpdateCustomerAddress(ByVal vXMLRequest As IXMLDOMElement)
' header -------------------------------------------------------------------------------
' description:
'   Update the persistant data associated with tables 'Customer Address', 'Address'.
'   If corresponding data exists in 'Current Property' update the record, else create
'   a record
' pass:
'   vXMLRequest  xml Request node containing data to be persisted
' Dependencies : CustomerBO, CustomerTxBO
' Raise Errors:
'---------------------------------------------------------------------------------------
On Error GoTo ICustomerDO_UpdateCustomerAddressVbErr

    Const strFunctionName As String = "ICustomerDO_UpdateCustomerAddress"

    Dim xmlElement As IXMLDOMElement
    Dim xmlValueElem As IXMLDOMElement
    Dim objErrAssist As New ErrAssist
    Dim objXmlAssist As New XMLAssist
    Dim objSQLAssist As New SQLAssist

    Dim objIDOAssist As IDOAssist
    Dim objIADOAssist  As IADOAssist
    Dim objIAddressDO As IAddressDO
    Dim objIClassDef As IomCustClassDef

    Dim strCustomerNumber As String
    Dim strCustomerVersionNumber As String
    Dim strCustAddrSeqNumber As String
    Dim strCondition As String

    Dim objContext As ObjectContext
    Set objContext = GetObjectContext()

    If objContext Is Nothing Then
        Set objIDOAssist = New DOAssist
        Set objIAddressDO = New AddressDO
        Set objIADOAssist = New ADOAssist
    Else
        Set objIDOAssist = objContext.CreateInstance(gstrBASE_COMPONENT & ".DOAssist")
        Set objIADOAssist = objContext.CreateInstance(gstrBASE_COMPONENT & ".ADOAssist")
        Set objIAddressDO = objContext.CreateInstance(cstrADDRESS_COMPONENT & ".AddressDO")
    End If

    'Fetch values of CustomerNumber, CustomerVersionNumber and CustomerAddessSequence No,
    'required for checking the existance of corresponding record in 'Current Property'
    If Not ICustomerDO_bGetKeysOfCustomerAddress(vXMLRequest, strCustomerNumber, _
                                     strCustomerVersionNumber, strCustAddrSeqNumber) Then
        objErrAssist.ThrowError TypeName(Me), strFunctionName, 519, _
                                            "for CUSTOMERADDRESS"
    End If

    'Check whether 'Address' data is to be modified; modify if required
    Set xmlElement = vXMLRequest.selectSingleNode("ADDRESS")
    If Not xmlElement Is Nothing Then
         objIAddressDO.UpdateAddress xmlElement
    End If
    
    Set objIClassDef = New omCustClassDef

    'Check whether 'CurrentProperty' data is to be modified; modify if required
    Set xmlElement = vXMLRequest.selectSingleNode("CURRENTPROPERTY")
    If Not xmlElement Is Nothing Then
        strCondition = "CUSTOMERNUMBER = " & objSQLAssist.FormatString(strCustomerNumber) _
                       & " And CUSTOMERVERSIONNUMBER = " & strCustomerVersionNumber _
                       & " And CUSTOMERADDRESSSEQUENCENUMBER = " & strCustAddrSeqNumber

        If objIADOAssist.CheckRecordExists("CURRENTPROPERTY", strCondition) Then
            objIDOAssist.Update xmlElement, objIClassDef.LoadCurrentPropertyData
        Else
            'Assign CustomerNumber, VersionNumber and AddressSeqNo to CurrentProperty XML
            Set xmlValueElem = xmlElement.getElementsByTagName("CUSTOMERNUMBER").Item(0)
            xmlValueElem.Text = strCustomerNumber
            Set xmlValueElem = xmlElement.getElementsByTagName("CUSTOMERVERSIONNUMBER").Item(0)
            xmlValueElem.Text = strCustomerVersionNumber
            Set xmlValueElem = xmlElement.getElementsByTagName("CUSTOMERADDRESSSEQUENCENUMBER").Item(0)
            xmlValueElem.Text = strCustAddrSeqNumber

            objIDOAssist.Create xmlElement, objIClassDef.LoadCurrentPropertyData
        End If
    End If

    ' SAB - 21/04/2006 - EP399 - Begin
    Dim xmlLastAmendedDate As IXMLDOMNode
    Set xmlLastAmendedDate = vXMLRequest.ownerDocument.createElement("LASTAMENDEDDATE")
    xmlLastAmendedDate.Text = Format(Now, "dd/mm/yyyy hh:mm:ss")
    vXMLRequest.appendChild xmlLastAmendedDate
    ' SAB - 21/04/2006 - EP399 - End

    objIDOAssist.Update vXMLRequest, objIClassDef.LoadCustomerAddressData
    
    Set xmlLastAmendedDate = Nothing ' EP399
    
    Set xmlElement = Nothing
    Set xmlValueElem = Nothing

    Set objIDOAssist = Nothing
    Set objIADOAssist = Nothing
    Set objIAddressDO = Nothing
    Set objIClassDef = Nothing

    Set objErrAssist = Nothing
    Set objXmlAssist = Nothing
    Set objSQLAssist = Nothing
    If Not objContext Is Nothing Then
        objContext.SetComplete
        Set objContext = Nothing
    End If

    Exit Sub
ICustomerDO_UpdateCustomerAddressVbErr:

    Set xmlLastAmendedDate = Nothing ' EP399
    
    Set xmlElement = Nothing
    Set xmlValueElem = Nothing

    Set objIDOAssist = Nothing
    Set objIADOAssist = Nothing
    Set objIAddressDO = Nothing
    Set objIClassDef = Nothing

    Set objErrAssist = Nothing
    Set objXmlAssist = Nothing
    Set objSQLAssist = Nothing

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
        Set objContext = Nothing
    End If

    Set objErrAssist = Nothing
    '   re-raise error for business object to interpret as appropriate
    Err.Raise Err.Number, Err.Source, Err.Description

End Sub

Public Sub ICustomerDO_DeleteCustomerAddress(ByVal vXMLRequest As IXMLDOMElement, _
                                Optional ByVal blnDeleteMortgageAccount As Boolean = False)
' header -------------------------------------------------------------------------------
' description:
'   Delete the persistant data associated with table 'Customer Address'.
'   If corresponding data exists in 'Current Property' Delete that record also.
' pass:
'   vXMLRequest  xml Request node containing data to be persisted
' Dependencies : CustomerBO, CustomerTxBO
' Raise Errors:
'---------------------------------------------------------------------------------------
On Error GoTo ICustomerDO_DeleteCustomerAddressVbErr

    Const strFunctionName As String = "ICustomerDO_DeleteCustomerAddress"

    Dim objSQLAssist As New SQLAssist

    Dim xmlCurrPropElem As IXMLDOMElement
    Dim xmlNodeList As IXMLDOMNodeList
    Dim xmlNode As IXMLDOMNode

    Dim objIADOAssist  As IADOAssist
    Dim objIAddressDO As IAddressDO
    Dim objICFDO As ICustomerFinancialDO
    Dim objIClassDef As IomCustClassDef

    Dim strCustomerNumber As String
    Dim strCustomerVersionNumber As String
    Dim strCustAddrSeqNumber As String
    Dim strCondition As String

    ' TO find the linked records in MortgageAccount and Delete those
    Dim xmlDoc As New FreeThreadedDOMDocument40
    xmlDoc.validateOnParse = False
    xmlDoc.setProperty "NewParser", True
    'Dim xmlMortgageList As IXMLDOMNode
    Dim xmlMortgageAccount  As IXMLDOMNode
    'BMIDS00006 Use AddressGUID as the key
    'Dim xmlCustomerNumber As IXMLDOMNode, xmlCustomerVersionNumber As IXMLDOMNode
    'Dim xmlCustAddressSeqNumber As IXMLDOMNode
    Dim xmlAddressGuid  As IXMLDOMNode
    Dim strAddressGuid  As String
    'BMIDS00006 End
    Dim xmlTempRespone As IXMLDOMNode
    Dim xmlMortgageAccountList As IXMLDOMNodeList
    Dim xmlAccountGuid As IXMLDOMNode, xmlThirdPartyGuid As IXMLDOMNode

    If m_objContext Is Nothing Then
        Set objIADOAssist = New ADOAssist
        Set objICFDO = New CustomerFinancialDO
    Else
        Set objIADOAssist = m_objContext.CreateInstance(gstrBASE_COMPONENT & ".ADOAssist")
        Set objICFDO = m_objContext.CreateInstance(gstrCUSTOMER_FINANCIAL & ".CustomerFinancialDO")
    End If

    'Fetch values of CustomerNumber, CustomerVersionNumber and CustomerAddessSequence No,
    'required for checking the existance of corresponding record in 'Current Property'
    If Not ICustomerDO_bGetKeysOfCustomerAddress(vXMLRequest, strCustomerNumber, _
                                    strCustomerVersionNumber, strCustAddrSeqNumber) Then
        m_objErrAssist.ThrowError TypeName(Me), strFunctionName, 519, _
                            "for CUSTOMERADDRESS"
    End If

    'BMIDS00006
    strAddressGuid = m_objXmlAssist.GetTagValue(vXMLRequest, "ADDRESSGUID")
    'BMIDS00006 End

    'Check whether 'CurrentProperty' data exist for this Customer Address;  If yes, delete
    strCondition = "CUSTOMERNUMBER = " & objSQLAssist.FormatString(strCustomerNumber) _
                   & " And CUSTOMERVERSIONNUMBER = " & strCustomerVersionNumber _
                   & " And CUSTOMERADDRESSSEQUENCENUMBER = " & strCustAddrSeqNumber
    
    Set objIClassDef = New omCustClassDef

    If objIADOAssist.CheckRecordExists("CURRENTPROPERTY", strCondition) Then

        'Build XML node to be passed for deleting record from table 'CurrentProperty' and
        ' add it to Request node
        Set xmlCurrPropElem = vXMLRequest.ownerDocument.createElement("CURRENTPROPERTY")
        Set xmlNode = vXMLRequest.appendChild(xmlCurrPropElem)
        Set xmlNode = vXMLRequest.selectSingleNode("//CUSTOMERNUMBER")
        Set xmlNode = xmlNode.cloneNode(True)
        Call xmlCurrPropElem.appendChild(xmlNode)

        Set xmlNode = vXMLRequest.selectSingleNode("//CUSTOMERVERSIONNUMBER")
        Set xmlNode = xmlNode.cloneNode(True)
        Call xmlCurrPropElem.appendChild(xmlNode)

        Set xmlNode = vXMLRequest.selectSingleNode("//CUSTOMERADDRESSSEQUENCENUMBER")
        Set xmlNode = xmlNode.cloneNode(True)
        Call xmlCurrPropElem.appendChild(xmlNode)

        m_objIDOAssist.Delete xmlCurrPropElem, objIClassDef.LoadCurrentPropertyData
    End If

    ' SR 01/09/00 : SYS1454 - For Homezone, delete the MortgageAccount records linked
    '               to the CustomerAddress
    If blnDeleteMortgageAccount Then
        Set xmlMortgageAccount = xmlDoc.createElement("MORTGAGEACCOUNT")
        
        ' BMIDS00006 Use account GUID as key for deleting mortgage accounts
        'Set xmlCustomerNumber = xmlDoc.createElement("CUSTOMERNUMBER")
        'xmlMortgageAccount.appendChild xmlCustomerNumber
        'xmlCustomerNumber.Text = strCustomerNumber
        '
        'Set xmlCustomerVersionNumber = xmlDoc.createElement("CUSTOMERVERSIONNUMBER")
        'xmlMortgageAccount.appendChild xmlCustomerVersionNumber
        'xmlCustomerVersionNumber.Text = strCustomerVersionNumber
        '
        'Set xmlCustAddressSeqNumber = xmlDoc.createElement("CUSTOMERADDRESSSEQUENCENUMBER")
        'xmlMortgageAccount.appendChild xmlCustAddressSeqNumber
        'xmlCustAddressSeqNumber.Text = strCustAddrSeqNumber
        
        'Find the AddressGuid required deletion of record from table 'Address'. First check
        'whether it is mentioned in input else get it from the table 'CustomerAddress'
        
        Set xmlAddressGuid = xmlDoc.createElement("ADDRESSGUID")
        xmlMortgageAccount.appendChild xmlAddressGuid
        xmlAddressGuid.Text = strAddressGuid
        'BMIDS00006 End
        
        On Error GoTo RecordNotFoundErr
        
        Set xmlTempRespone = objICFDO.FindMortgageAccountList(xmlMortgageAccount)
        
        On Error GoTo ICustomerDO_DeleteCustomerAddressVbErr
        If Not xmlTempRespone Is Nothing Then
            Set xmlMortgageAccountList = xmlTempRespone.selectNodes(".//MORTGAGEACCOUNT")
            If xmlMortgageAccountList.length > 0 Then
                Set xmlAccountGuid = xmlDoc.createElement("ACCOUNTGUID")
                xmlMortgageAccount.appendChild xmlAccountGuid
            
                Set xmlThirdPartyGuid = xmlDoc.createElement("THIRDPARTYGUID")
                xmlMortgageAccount.appendChild xmlThirdPartyGuid
            End If
            
            Dim i As Long
            For i = 0 To xmlMortgageAccountList.length - 1
                Set xmlNode = xmlMortgageAccountList(i)
                xmlAccountGuid.Text = xmlNode.selectSingleNode(".//ACCOUNTGUID").Text
                xmlThirdPartyGuid.Text = xmlNode.selectSingleNode(".//THIRDPARTYGUID").Text
                
                objICFDO.DeleteMortgageAccount xmlMortgageAccount
            Next i
        End If
    Else
        'SR - AQR/SYS0392 : if any of the records in 'MortgageAccount' are referring to
        ' this customeradrdres, then set the CustomeAddressSequenceNumber to NULL for
        ' in MortgageAccount records.
        Dim strSQL As String
    
        'BMIDS00006
        'strSQL = " UPDATE MORTGAGEACCOUNT SET CUSTOMERADDRESSSEQUENCENUMBER=NULL WHERE " _
        '            & strCondition
        strSQL = " UPDATE MORTGAGEACCOUNT SET ADDRESSGUID=NULL WHERE ADDRESSGUID='" & _
            objSQLAssist.FormatGuid(strAddressGuid) & "'"
        'BMIDS00006 End
    
        On Error Resume Next
        objIADOAssist.executeSQLCommand strSQL, True
        m_objErrAssist.SaveErr
    
        On Error GoTo ICustomerDO_DeleteCustomerAddressVbErr
        If m_objErrAssist.LastErr <> 0 And _
            m_objErrAssist.GetOmigaErrorNumber(m_objErrAssist.LastErr) <> oeNoRowsAffected Then
            m_objErrAssist.ReRaise
        End If
    End If
    
    '  Delete the CustomerAddress data
    m_objIDOAssist.Delete vXMLRequest, objIClassDef.LoadCustomerAddressData

    Set xmlCurrPropElem = Nothing
    Set xmlNode = Nothing

    Set objIADOAssist = Nothing
    Set objIClassDef = Nothing
    Set objIAddressDO = Nothing
    Set objICFDO = Nothing

    Set xmlMortgageAccount = Nothing
    'Set xmlCustomerNumber = Nothing
    'Set xmlCustomerVersionNumber = Nothing
    'Set xmlCustAddressSeqNumber = Nothing
    Set xmlAddressGuid = Nothing
    Set xmlTempRespone = Nothing
    Set xmlMortgageAccountList = Nothing
    Set xmlAccountGuid = Nothing
    Set xmlThirdPartyGuid = Nothing

    Set objSQLAssist = Nothing
    
    If Not m_objContext Is Nothing Then
        m_objContext.SetComplete
    End If

    Exit Sub
RecordNotFoundErr:
    
    If m_objErrAssist.GetOmigaErrorNumber(Err.Number) = oeRecordNotFound Then
        Resume Next
    End If

ICustomerDO_DeleteCustomerAddressVbErr:

    Set xmlCurrPropElem = Nothing
    Set xmlNode = Nothing

    Set objIADOAssist = Nothing
    Set objIClassDef = Nothing
    Set objIAddressDO = Nothing
    
    Set objSQLAssist = Nothing

    m_objErrAssist.AddToErrSource (strFunctionName)

    If m_objErrAssist.IsSystemError = True Then
        m_objErrAssist.LogError TypeName(Me), strFunctionName, Err.Number, Err.Description
    End If

    If Not m_objContext Is Nothing Then
        If m_objErrAssist.IsSystemError Then
            m_objContext.SetAbort
        Else
            m_objContext.SetComplete
        End If
    End If

    '   re-raise error for business object to interpret as appropriate
    Err.Raise Err.Number, Err.Source, Err.Description

End Sub

Private Function ICustomerDO_bGetKeysOfCustomerAddress(ByVal xmlElement As IXMLDOMElement, _
                                          ByRef strCustomerNumber As String, _
                                          ByRef strCustomerVersionNumber As String, _
                                          ByRef strCustomerAddressSequenceNumber As String) As Boolean
' header -------------------------------------------------------------------------------
' description:
'   Fetches the values for CustomerNumber, CustomerVersionNumber, CustomerAddressSequenceNumber
'   from xmlElement.
' pass:
'   Returns False, if any of the values is empty. Otherwise True
' Dependencies : CustomerBO, CustomerTxBO
' Raise Errors:
'---------------------------------------------------------------------------------------
On Error GoTo ICustomerDO_bGetKeysOfCustomerAddressVbErr
    Const strFunctionName As String = "ICustomerDO_bGetKeysOfCustomerAddress"

    Dim objXmlAssist As New XMLAssist
    Dim objErrAssist As New ErrAssist

    strCustomerNumber = objXmlAssist.GetTagValue(xmlElement, "CUSTOMERNUMBER")
    strCustomerVersionNumber = objXmlAssist.GetTagValue(xmlElement, "CUSTOMERVERSIONNUMBER")
    strCustomerAddressSequenceNumber = objXmlAssist.GetTagValue(xmlElement, _
                                                    "CUSTOMERADDRESSSEQUENCENUMBER")

    If strCustomerNumber = "" Or strCustomerVersionNumber = "" _
                             Or strCustomerAddressSequenceNumber = "" Then
        ICustomerDO_bGetKeysOfCustomerAddress = False
    Else
        ICustomerDO_bGetKeysOfCustomerAddress = True
    End If

    Set objXmlAssist = Nothing
    Set objErrAssist = Nothing

    Exit Function
ICustomerDO_bGetKeysOfCustomerAddressVbErr:

    Set objXmlAssist = Nothing
    objErrAssist.AddToErrSource (strFunctionName)

    If objErrAssist.IsSystemError = True Then
        objErrAssist.LogError TypeName(Me), strFunctionName, Err.Number, Err.Description
    End If

    Set objErrAssist = Nothing

    're-raise error for business object to interpret as appropriate
    Err.Raise Err.Number, Err.Source, Err.Description
End Function

Private Sub ICustomerDO_CreateTenancy(ByVal vxmlTableElement As IXMLDOMElement)
' header ----------------------------------------------------------------------------------
' description:
'   CreateTenancy an instance of the persistant data associated with this data object
'   for each set of data in the request
' pass:
'   vxmlxmlTableElement  xml element containing the request
' return:                n/a
'------------------------------------------------------------------------------------------
On Error GoTo CreateTenancyVbErr

    Const strFunctionName As String = "ICustomerDO_CreateTenancy"
   
    Dim objErrAssist As New ErrAssist
    Dim xmlClassDefDoc As FreeThreadedDOMDocument40
    Dim objClassDef As IomCustClassDef

    Dim objXmlAssist As New XMLAssist, _
        objTPAssist As IThirdPartyAssist
    Set objTPAssist = New ThirdPartyAssist

    
    Dim objIDOAssist As IDOAssist
    
    Dim objContext As ObjectContext
    Set objContext = GetObjectContext()
                    
    If objContext Is Nothing Then
        Set objIDOAssist = New DOAssist
    Else
        Set objIDOAssist = objContext.CreateInstance(gstrBASE_COMPONENT & ".DOAssist")
    End If

    '
    ' Maintain THIRDPARTY if necessary
    '
    If (Trim$(objXmlAssist.GetTagValue(vxmlTableElement, "THIRDPARTY")) <> "") And _
       (Trim$(objXmlAssist.GetTagValue(vxmlTableElement, "DIRECTORYGUID")) = "") Then
        ' THIRDPARTY details exist in the XML, so maintain the appropriate THIRDPARTY record
        objTPAssist.MaintainThirdParty vxmlTableElement
    End If

    '
    ' Maintain main table
    '
    Set objClassDef = New omCustClassDef
    Set xmlClassDefDoc = objClassDef.LoadTenancyData
    objIDOAssist.Create vxmlTableElement, xmlClassDefDoc
    
    Set objClassDef = Nothing
    Set xmlClassDefDoc = Nothing
    Set objIDOAssist = Nothing
    Set objErrAssist = Nothing
    Set objXmlAssist = Nothing
    Set objTPAssist = Nothing

    If Not objContext Is Nothing Then
        objContext.SetComplete
    End If

    Exit Sub

CreateTenancyVbErr:
    
    Set objClassDef = Nothing
    Set xmlClassDefDoc = Nothing
    Set objIDOAssist = Nothing
    Set objXmlAssist = Nothing
    Set objTPAssist = Nothing
    
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

Private Function ICustomerDO_UpdateCustomerVersion( _
        ByVal vxmlTableElement As IXMLDOMElement, _
        Optional ByVal bTaxDetailsOnly As Boolean = False) As IXMLDOMElement

On Error GoTo UpdateVbErr

    Const strFunctionName As String = "ICustomerDO_UpdateCustomerVersion"
                
    Dim objContext As ObjectContext
    Set objContext = GetObjectContext()
    
    Dim objIClassDef As IomCustClassDef
    Dim xmlClassDefDoc As FreeThreadedDOMDocument40
    Dim objIDOAssist As IDOAssist
    Dim objErrAssist As New ErrAssist
    
    Set objIClassDef = New omCustClassDef
    
    If bTaxDetailsOnly Then
        Set xmlClassDefDoc = objIClassDef.LoadCustomerVersionData(omCustVersionTax)
    Else
        Set xmlClassDefDoc = objIClassDef.LoadCustomerVersionData
    End If
        
    If objContext Is Nothing Then
        Set objIDOAssist = New DOAssist
    Else
        Set objIDOAssist = objContext.CreateInstance(gstrBASE_COMPONENT & ".DOAssist")
    End If
        
    objIDOAssist.Update vxmlTableElement, xmlClassDefDoc
        
    Set objIClassDef = Nothing
    Set xmlClassDefDoc = Nothing
    Set objIDOAssist = Nothing
    Set objErrAssist = Nothing
    
    If Not objContext Is Nothing Then
        objContext.SetComplete
    End If
    
    Exit Function

UpdateVbErr:
        
    Set objIClassDef = Nothing
    Set xmlClassDefDoc = Nothing
    Set objIDOAssist = Nothing
    
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

    '   re-raise error for business object to interpret as appropriate
    Err.Raise Err.Number, Err.Source, Err.Description

End Function

Private Sub ICustomerDO_UpdateHomeInsurance(ByVal vxmlTableElement As IXMLDOMElement)
' header ----------------------------------------------------------------------------------
' description:
'   Update a single instance of the persistant data associated with this
'   data object on the HomeInsurance Table
' pass:
'   vxmlTableElement  xml element containing the request
' return:                n/a
'------------------------------------------------------------------------------------------
On Error GoTo UpdateHomeInsuranceVbErr

    Const strFunctionName As String = "ICustomerDO_UpdateHomeInsurance"
            
    Dim xmlClassDefDoc As FreeThreadedDOMDocument40
    Dim objIClassDef As IomCustClassDef
    
  
    Set objIClassDef = New omCustClassDef
    Set xmlClassDefDoc = objIClassDef.LoadHomeInsuranceData()
    m_objIDOAssist.Update vxmlTableElement, xmlClassDefDoc
    
    Set objIClassDef = Nothing
    Set xmlClassDefDoc = Nothing
    
    If Not m_objContext Is Nothing Then
        m_objContext.SetComplete
    End If
    
    Exit Sub

UpdateHomeInsuranceVbErr:
        
    Set objIClassDef = Nothing
    Set xmlClassDefDoc = Nothing
    
    m_objErrAssist.AddToErrSource (strFunctionName)
    
    If m_objErrAssist.IsSystemError = True Then
        m_objErrAssist.LogError TypeName(Me), strFunctionName, Err.Number, Err.Description
    End If
    
    If Not m_objContext Is Nothing Then
        If m_objErrAssist.IsSystemError Then
            m_objContext.SetAbort
        Else
            m_objContext.SetComplete
        End If
    End If

    '   re-raise error for business object to interpret as appropriate
    Err.Raise Err.Number, Err.Source, Err.Description

End Sub

Private Sub ICustomerDO_UpdatePropertyInsuranceDetails(ByVal vxmlTableElement As IXMLDOMElement)
' header ----------------------------------------------------------------------------------
' description:
'   Update a single instance of the persistant data associated with this
'   data object by calling UpdateCurrentProperty and UpdateHomeInsurance
' pass:
'   vxmlTableElement  xml element containing the request
' return:                n/a
'------------------------------------------------------------------------------------------
On Error GoTo UpdatePropertyInsuranceDetailsVbErr

    Const strFunctionName As String = "ICustomerDO_UpdatePropertyInsuranceDetails"
            
    ICustomerDO_UpdateCurrentProperty vxmlTableElement
    ICustomerDO_UpdateHomeInsurance vxmlTableElement
    
    If Not m_objContext Is Nothing Then
        m_objContext.SetComplete
    End If
    
    Exit Sub

UpdatePropertyInsuranceDetailsVbErr:
        
 
    m_objErrAssist.AddToErrSource (strFunctionName)
    
    If m_objErrAssist.IsSystemError = True Then
        m_objErrAssist.LogError TypeName(Me), strFunctionName, Err.Number, Err.Description
    End If
    
    If Not m_objContext Is Nothing Then
        If m_objErrAssist.IsSystemError Then
            m_objContext.SetAbort
        Else
            m_objContext.SetComplete
        End If
    End If

    '   re-raise error for business object to interpret as appropriate
    Err.Raise Err.Number, Err.Source, Err.Description

End Sub

Private Sub ICustomerDO_UpdateTenancy(ByVal vxmlTableElement As IXMLDOMElement)
' header ----------------------------------------------------------------------------------
' description:
'   UpdateTenancy a single instance of the persistant data associated with this
'   data object
' pass:
'   vxmlxmlTableElement  xml element containing the request
' return:                n/a
'------------------------------------------------------------------------------------------
On Error GoTo UpdateTenancyVbErr

    Const strFunctionName As String = "ICustomerDO_UpdateTenancy"
        
    Dim objErrAssist As New ErrAssist
    
    Dim objContext As ObjectContext
    Set objContext = GetObjectContext()
    
    Dim xmlClassDefDoc As FreeThreadedDOMDocument40
    Dim objClassDef As IomCustClassDef
    Dim objIDOAssist As IDOAssist, _
        objXmlAssist As New XMLAssist, _
        objTPAssist As IThirdPartyAssist
    Set objTPAssist = New ThirdPartyAssist

    Dim xmlThirdPartyGUIDNode As IXMLDOMNode
    Dim strThirdPartyGUID As String, _
        strDirectoryGUID  As String

    If objContext Is Nothing Then
        Set objIDOAssist = New DOAssist
    Else
        Set objIDOAssist = objContext.CreateInstance(gstrBASE_COMPONENT & ".DOAssist")
    End If

    Set objClassDef = New omCustClassDef
    Set xmlClassDefDoc = objClassDef.LoadTenancyData()

    '
    ' Maintain the associated THIRDPARTY record if appropriate
    '
    strThirdPartyGUID = Trim$(objXmlAssist.GetTagValue(vxmlTableElement, "THIRDPARTYGUID"))
    strDirectoryGUID = Trim$(objXmlAssist.GetTagValue(vxmlTableElement, "DIRECTORYGUID"))

    If (strThirdPartyGUID <> "") And (strDirectoryGUID <> "") Then
        ' Cannot have both a DIRECTORY and THIRDPARTY GUID. Remove the THIRDPARTYGUID
        Set xmlThirdPartyGUIDNode = vxmlTableElement.selectSingleNode("THIRDPARTYGUID")
        If Not (xmlThirdPartyGUIDNode Is Nothing) Then
            xmlThirdPartyGUIDNode.Text = ""
        End If

        ' Update the primary table with the null THIRDPARTYGUID
        objIDOAssist.Update vxmlTableElement, xmlClassDefDoc
        ' Delete from the THIRDPARTY table
        objTPAssist.DeleteThirdParty strThirdPartyGUID
    Else
        objTPAssist.MaintainThirdParty vxmlTableElement
        objIDOAssist.Update vxmlTableElement, xmlClassDefDoc
    End If
    
    Set objClassDef = Nothing
    Set xmlClassDefDoc = Nothing
    Set objIDOAssist = Nothing
    Set objErrAssist = Nothing
    Set objXmlAssist = Nothing
    Set objTPAssist = Nothing
    
    If Not objContext Is Nothing Then
        objContext.SetComplete
    End If
    
    Exit Sub

UpdateTenancyVbErr:
        
    Set objClassDef = Nothing
    Set xmlClassDefDoc = Nothing
    Set objIDOAssist = Nothing
    Set objXmlAssist = Nothing
    Set objTPAssist = Nothing
    
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

Private Sub ICustomerDO_DeleteTenancy(ByVal vxmlTableElement As IXMLDOMElement)
' header ----------------------------------------------------------------------------------
' description:
'   DeleteTenancy a single instance of the persistant data associated with this
'   data object
' pass:
'   vxmlxmlTableElement  xml element containing the request
' return:                n/a
'------------------------------------------------------------------------------------------
On Error GoTo DeleteTenancyVbErr

    Const strFunctionName As String = "ICustomerDO_DeleteTenancy"

    Dim objErrAssist As New ErrAssist
    
    Dim objContext As ObjectContext
    Set objContext = GetObjectContext()
            
    Dim objIDOAssist As IDOAssist, _
        objXmlAssist As New XMLAssist, _
        objTPAssist As IThirdPartyAssist
    Set objTPAssist = New ThirdPartyAssist

    Dim xmlClassDefDoc As FreeThreadedDOMDocument40
    Dim objClassDef As IomCustClassDef

    Dim xmlData As IXMLDOMNode

    Dim strThirdPartyGUID As String
    
    If objContext Is Nothing Then
        Set objIDOAssist = New DOAssist
    Else
        Set objIDOAssist = objContext.CreateInstance(gstrBASE_COMPONENT & ".DOAssist")
    End If
            
    Set objClassDef = New omCustClassDef
    Set xmlClassDefDoc = objClassDef.LoadTenancyData()

    ' Retrieve THIRDPARTYGUID
    Set xmlData = ICustomerDO_GetTenancy(vxmlTableElement)
    strThirdPartyGUID = objXmlAssist.GetTagValue(xmlData, "THIRDPARTYGUID")

    objIDOAssist.Delete vxmlTableElement, xmlClassDefDoc
    
    If strThirdPartyGUID <> "" Then
        ' Delete the thirdparty details
        objTPAssist.DeleteThirdParty strThirdPartyGUID
    End If

    Set objClassDef = Nothing
    Set xmlClassDefDoc = Nothing
    Set objIDOAssist = Nothing
    Set objErrAssist = Nothing
    Set objXmlAssist = Nothing
    Set objTPAssist = Nothing
    
    If Not objContext Is Nothing Then
        objContext.SetComplete
    End If
    
    Exit Sub

DeleteTenancyVbErr:

    Set objClassDef = Nothing
    Set xmlClassDefDoc = Nothing
    Set objIDOAssist = Nothing
    Set objXmlAssist = Nothing
    Set objTPAssist = Nothing
    
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

Private Function ICustomerDO_GetTenancy(ByVal vxmlTableElement As IXMLDOMElement) As IXMLDOMNode
' header ----------------------------------------------------------------------------------
' description:
'   Get the data for a single instance of the persistant data associated with
'   this data object
' pass:
'   vxmlxmlTableElement  xml element containing the request
' return:                xml node containing retrieved data
'------------------------------------------------------------------------------------------
On Error GoTo GetTenancyVbErr

    Const strFunctionName As String = "ICustomerDO_GetTenancy"
    
    Dim objErrAssist As New ErrAssist
    
    Dim objIDOAssist As IDOAssist, _
        objTPAssist As IThirdPartyAssist
    Set objTPAssist = New ThirdPartyAssist

    Dim xmlClassDefDoc As FreeThreadedDOMDocument40
    Dim objClassDef As IomCustClassDef
    Dim xmlReturn As IXMLDOMNode
    
    Dim objContext As ObjectContext
    Set objContext = GetObjectContext()
            
    If objContext Is Nothing Then
        Set objIDOAssist = New DOAssist
    Else
        Set objIDOAssist = objContext.CreateInstance(gstrBASE_COMPONENT & ".DOAssist")
    End If
    
    Set objClassDef = New omCustClassDef
    Set xmlClassDefDoc = objClassDef.LoadTenancyData()
    Set xmlReturn = objIDOAssist.GetData(vxmlTableElement, xmlClassDefDoc)
    
    objTPAssist.GetThirdParty xmlReturn

    Set ICustomerDO_GetTenancy = xmlReturn

    Set objClassDef = Nothing
    Set xmlClassDefDoc = Nothing
    Set objIDOAssist = Nothing
    Set objErrAssist = Nothing
    Set xmlReturn = Nothing
    Set objTPAssist = Nothing
    
    If Not objContext Is Nothing Then
        objContext.SetComplete
    End If
    
    Exit Function

GetTenancyVbErr:
    
    Set objClassDef = Nothing
    Set xmlClassDefDoc = Nothing
    Set objIDOAssist = Nothing
    Set xmlReturn = Nothing
    Set objTPAssist = Nothing
    
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
End Function

Private Sub ICustomerDO_CreateCustomerTelephone(ByVal vXMLRequest As IXMLDOMElement)
                                                    
' header ----------------------------------------------------------------------------------
' description:
'   create an instance of the persistant data associated with this data object
'   for each set of data in the request
' pass:
'   vXMLRequest   xml Request node containing data to be persisted
' Raise Errors:
'------------------------------------------------------------------------------------------
On Error GoTo ICustomerDO_CreateCustomerTelephoneVbErr

    Const strFunctionName As String = "ICustomerDO_CreateCustomerTelephone"

    Dim objErrAssist As New ErrAssist
    Dim objXmlAssist As New XMLAssist
    Dim xmlClassDef As FreeThreadedDOMDocument40

    Dim objIDOAssist As IDOAssist
    Dim objIClassDef As IomCustClassDef

    Dim objContext As ObjectContext
    Set objContext = GetObjectContext()

    If objContext Is Nothing Then
        Set objIDOAssist = New DOAssist
    Else
        Set objIDOAssist = objContext.CreateInstance(gstrBASE_COMPONENT & ".DOAssist")
    End If

    '
    ' Insert the wrapup telephone sequence number
    '
    Set objIClassDef = New omCustClassDef
    Set xmlClassDef = objIClassDef.LoadCustomerTelephoneData
    objIDOAssist.GenerateSequenceNumber vXMLRequest, xmlClassDef, _
                                         "TELEPHONESEQUENCENUMBER"

    objIDOAssist.Create vXMLRequest, xmlClassDef

    Set objIDOAssist = Nothing
    Set objErrAssist = Nothing
    Set xmlClassDef = Nothing
    Set objXmlAssist = Nothing

    If Not objContext Is Nothing Then
        objContext.SetComplete
        Set objContext = Nothing
    End If

    Exit Sub

ICustomerDO_CreateCustomerTelephoneVbErr:

    Set objIDOAssist = Nothing
    Set objIClassDef = Nothing
    Set xmlClassDef = Nothing
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
        Set objContext = Nothing
    End If

    Set objErrAssist = Nothing

    '   re-raise error for business object to interpret as appropriate
    Err.Raise Err.Number, Err.Source, Err.Description
End Sub

Private Sub ICustomerDO_CreateWrapUpDetails(ByVal vXMLRequest As IXMLDOMElement)
                                                    
' header ----------------------------------------------------------------------------------
' description:
'   create an instance of the persistant data associated with this data object
'   for each set of data in the request
' pass:
'   vXMLRequest   xml Request node containing data to be persisted
' Raise Errors:
'------------------------------------------------------------------------------------------
On Error GoTo ICustomerDO_CreateWrapUpDetailsVbErr

    Const strFunctionName As String = "ICustomerDO_CreateWrapUpDetails"

    Dim objErrAssist As New ErrAssist
    Dim objXmlAssist As New XMLAssist
    Dim xmlClassDef As FreeThreadedDOMDocument40

    Dim objIDOAssist As IDOAssist
    Dim objIClassDef As IomCustClassDef

    Dim objContext As ObjectContext
    Set objContext = GetObjectContext()

    If objContext Is Nothing Then
        Set objIDOAssist = New DOAssist
    Else
        Set objIDOAssist = objContext.CreateInstance(gstrBASE_COMPONENT & ".DOAssist")
    End If

    '
    ' Insert the customer telephone sequence number
    '
    Set objIClassDef = New omCustClassDef
    Set xmlClassDef = objIClassDef.LoadWrapUpDetailsData
    objIDOAssist.GenerateSequenceNumber vXMLRequest, xmlClassDef, _
                                         "WRAPUPSEQUENCENUMBER"

    objIDOAssist.Create vXMLRequest, xmlClassDef

    Set objIDOAssist = Nothing
    Set objErrAssist = Nothing
    Set xmlClassDef = Nothing
    Set objXmlAssist = Nothing

    If Not objContext Is Nothing Then
        objContext.SetComplete
        Set objContext = Nothing
    End If

    Exit Sub

ICustomerDO_CreateWrapUpDetailsVbErr:

    Set objIDOAssist = Nothing
    Set objIClassDef = Nothing
    Set xmlClassDef = Nothing
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
        Set objContext = Nothing
    End If

    Set objErrAssist = Nothing

    '   re-raise error for business object to interpret as appropriate
    Err.Raise Err.Number, Err.Source, Err.Description
End Sub

Private Function ICustomerDO_DeleteDependant(ByVal vXMLRequest As IXMLDOMElement, _
                           Optional ByRef rstrPersonGUID As Variant) As IXMLDOMElement
' header ----------------------------------------------------------------------------------
' description:
'   Delete a single instance of the persistant data associated with this
'   data object
' pass:
'   vXMLRequest  xml Request node containing data to which identifies
'                   the instance to be deleted
'   rstrPersonGUID : GUID of the person associated with this Dependant
' Raise Errors:
'       parser errors
'------------------------------------------------------------------------------------------
On Error GoTo ICustomerDO_DeleteDependantVbErr

    Const strFunctionName As String = "ICustomerDO_DeleteDependant"

    Dim objErrAssist As New ErrAssist

    Dim objContext As ObjectContext
    Set objContext = GetObjectContext()

    Dim objIDOAssist As IDOAssist
    Dim objXmlAssist As New XMLAssist

    Dim xmlPersonNode As IXMLDOMNode
    Dim objIPersonDO As IPersonDO
    Dim objIClassDef As IomCustClassDef

    Dim strPersonGuid As String
    Dim xmlElement As IXMLDOMElement

    '
    ' Initialise
    '
    If objContext Is Nothing Then
        Set objIDOAssist = New DOAssist
        Set objIPersonDO = New PersonDO
    Else
        Set objIDOAssist = objContext.CreateInstance(gstrBASE_COMPONENT & ".DOAssist")
        Set objIPersonDO = objContext.CreateInstance(cstrPERSON_COMPONENT & ".PersonDO")
    End If

    strPersonGuid = ""

    '
    ' Main
    ' If the Person Data is mentioned in the Request, use it for deletion
    ' Else call GetDependant details for the corresponding Person data
    Set xmlElement = vXMLRequest.getElementsByTagName("PERSON").Item(0)
    If xmlElement Is Nothing Then
        ' Retrieve the full details for the dependant so that the Person information can be ascertained
        Set xmlElement = ICustomerDO_GetDependant(vXMLRequest)
    ElseIf Len(objXmlAssist.GetTagValue(xmlElement, "PERSONGUID")) = 0 Then
        Set xmlElement = ICustomerDO_GetDependant(vXMLRequest)
    End If

    ' Delete the dependant
    Set objIClassDef = New omCustClassDef
    objIDOAssist.Delete vXMLRequest, objIClassDef.LoadDependantData

    ' Delete the corresponding Person data
    If xmlElement.nodeName = "PERSON" Then
        Set xmlPersonNode = xmlElement
    Else
        Set xmlPersonNode = xmlElement.selectSingleNode("//PERSON")
    End If
    
    strPersonGuid = objXmlAssist.GetTagValue(xmlPersonNode, "PERSONGUID")
    If strPersonGuid <> "" Then
        Call objIPersonDO.Delete(xmlPersonNode)
    End If

    ' Set value of by out parameter
    If Not IsMissing(rstrPersonGUID) Then
        rstrPersonGUID = strPersonGuid
    End If

    Set objIDOAssist = Nothing
    Set objXmlAssist = Nothing
    Set objErrAssist = Nothing
    Set xmlElement = Nothing
    Set objIPersonDO = Nothing

    If Not objContext Is Nothing Then
        objContext.SetComplete
        Set objContext = Nothing
    End If

    Exit Function

ICustomerDO_DeleteDependantVbErr:

    If (objErrAssist.GetOmigaErrorNumber(Err.Number) = oeRecordNotFound) Or _
       (objErrAssist.GetOmigaErrorNumber(Err.Number) = oeChildRecordsFound) Then
        Resume Next
    End If

    Set objIDOAssist = Nothing
    Set objXmlAssist = Nothing
    Set objIPersonDO = Nothing
    Set xmlElement = Nothing

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
        Set objContext = Nothing
    End If

    Set objErrAssist = Nothing
    '   re-raise error for business object to interpret as appropriate
    Err.Raise Err.Number, Err.Source, Err.Description

End Function

Private Function ICustomerDO_GetDependant(ByVal vXMLRequest As IXMLDOMElement) As IXMLDOMElement
' header ----------------------------------------------------------------------------------
' description:
'   Get the data for a single instance of the persistant data associated with
'   this data object
' pass:
'   vXMLRequest  xml Request node containing data to which identifies
'                   the instance of the persistant data to be retrieved
' return:
'   GetDependant    node containing XML data retrieved
' Raise Errors: if record not found, raise TempRecordNotFound
'------------------------------------------------------------------------------------------
On Error GoTo ICustomerDO_GetDependantVbErr

    Const strFunctionName As String = "ICustomerDO_GetDependant"

    Dim objErrAssist As New ErrAssist

    Dim objIDOAssist As IDOAssist
    Dim objIClassDef As IomCustClassDef
    
    Dim xmlResponseElem As IXMLDOMElement
    Dim xmlDependantElem As IXMLDOMElement
    Dim xmlData As IXMLDOMNode
    Dim objXmlAssist As New XMLAssist
    Dim strXML As String
    Dim xmlRequestDoc As FreeThreadedDOMDocument40
    
    Dim objContext As ObjectContext
    Set objContext = GetObjectContext()

    If objContext Is Nothing Then
        Set objIDOAssist = New DOAssist
    Else
        Set objIDOAssist = objContext.CreateInstance(gstrBASE_COMPONENT & ".DOAssist")
    End If

    Set xmlResponseElem = vXMLRequest.ownerDocument.createElement("RESPONSE")
    xmlResponseElem.setAttribute "TYPE", "SUCCESS"
       
    strXML = vXMLRequest.xml
    strXML = Replace(strXML, "<DEPENDANT>", "<DEPENDANTDETAILS>")
    strXML = Replace(strXML, "</DEPENDANT>", "</DEPENDANTDETAILS>")
    
    Set xmlRequestDoc = objXmlAssist.load(strXML, TypeName(Me), strFunctionName)
    
    Set objIClassDef = New omCustClassDef
        
    ' Arrange the 'PERSON' node under 'DEPENDANT' node in the hierarchy
    Set xmlData = objIDOAssist.GetData(xmlRequestDoc.documentElement, objIClassDef.LoadDependantDetailsData)
    Set xmlDependantElem = xmlData.selectSingleNode("DEPENDANT")
    
    xmlDependantElem.appendChild xmlData.selectSingleNode("PERSON")
    
    Set ICustomerDO_GetDependant = xmlData.cloneNode(True)

    Set objIDOAssist = Nothing
    Set objErrAssist = Nothing

    If Not objContext Is Nothing Then
        objContext.SetComplete
    End If

    Set xmlResponseElem = Nothing
    Set xmlDependantElem = Nothing
    Set xmlRequestDoc = Nothing
    Set xmlData = Nothing
    Set objXmlAssist = Nothing
    Set objErrAssist = Nothing

    Exit Function

ICustomerDO_GetDependantVbErr:
    Set xmlResponseElem = Nothing
    Set xmlDependantElem = Nothing
    Set xmlRequestDoc = Nothing
    Set xmlData = Nothing
    Set objXmlAssist = Nothing

    Set objIDOAssist = Nothing

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

End Function

Private Function ICustomerDO_FindCustomerTelephoneList(ByVal vXMLRequest As IXMLDOMElement) As IXMLDOMElement
' header ----------------------------------------------------------------------------------
' description:
'   Get the data for all instances of the persistant data associated with
'   this data object for the values supplied
' pass:
'   vXMLRequest  xml Request node containing data to which identifies
'                the instance(s) of the persistant data to be retrieved
' return:
'   FindCustomerTelephoneList  Node with the data retreived from database
' Raise Errors:
'   if record not found, raise TempRecordNotFound
'------------------------------------------------------------------------------------------
On Error GoTo ICustomerDO_FindCustomerTelephoneListVbErr

    Const strFunctionName As String = "ICustomerDO_FindCustomerTelephoneList"

    Dim objErrAssist As New ErrAssist

    Dim objIDOAssist As IDOAssist
    Dim objIClassDef As IomCustClassDef

    Dim objContext As ObjectContext
    Set objContext = GetObjectContext()

    If objContext Is Nothing Then
        Set objIDOAssist = New DOAssist
    Else
        Set objIDOAssist = objContext.CreateInstance(gstrBASE_COMPONENT & ".DOAssist")
    End If

    Set objIClassDef = New omCustClassDef
    Set ICustomerDO_FindCustomerTelephoneList = _
            objIDOAssist.FindList(vXMLRequest, objIClassDef.LoadCustomerTelephoneData)

    Set objErrAssist = Nothing
    Set objIDOAssist = Nothing
    Set objIClassDef = Nothing
    
    If Not objContext Is Nothing Then
        objContext.SetComplete
        Set objContext = Nothing
    End If

    Exit Function

ICustomerDO_FindCustomerTelephoneListVbErr:

    Set objIDOAssist = Nothing
    Set objIClassDef = Nothing

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
        Set objContext = Nothing
    End If

    Set objErrAssist = Nothing

    '   re-raise error for business object to interpret as appropriate
    Err.Raise Err.Number, Err.Source, Err.Description

End Function

Private Function ICustomerDO_UpdateDependant(ByVal vXMLRequest As IXMLDOMElement) As IXMLDOMElement
' header ----------------------------------------------------------------------------------
' description:
'   Update a single instance of the persistant data associated with this
'   data object
' pass:
'   vstrXMLRequest  xml Request data stream containing data to be persisted
' Raise Errors:
'------------------------------------------------------------------------------------------
On Error GoTo ICustomerDO_UpdateDependantVbErr

    Const strFunctionName As String = "ICustomerDO_UpdateDependant"

    Dim objErrAssist As New ErrAssist

    Dim objContext As ObjectContext
    Set objContext = GetObjectContext()

    Dim objIDOAssist As IDOAssist

    Dim objXmlAssist As New XMLAssist
    Dim xmlNode As IXMLDOMNode
    Dim xmlElement As IXMLDOMElement

    Dim botPersonOperation As BOOPERATIONTYPE
    Dim strPersonGuid As String
    Dim objIClassDef As IomCustClassDef
    Dim objIPersonClassDef As IPersonClassDef

    If objContext Is Nothing Then
        Set objIDOAssist = New DOAssist
    Else
        Set objIDOAssist = objContext.CreateInstance(gstrBASE_COMPONENT & ".DOAssist")
    End If

    '
    ' Initialise
    '
    Set objIPersonClassDef = New PersonClassdef
    botPersonOperation = objXmlAssist.DetermineOperationEx(vXMLRequest, objIPersonClassDef.LoadPersonData)
    strPersonGuid = ""

    '
    ' Create/Update Person
    '
    If (botPersonOperation = booCreate) Or (botPersonOperation = booUpdate) Then
        ' Person should be created/update before the third party record is updated
        SavePerson vXMLRequest, strPersonGuid, vbotOperationType:=botPersonOperation
    End If

    '
    ' Save Dependant
    '
    ' Insert Person GUID into dependant XML
    Set xmlNode = vXMLRequest.selectSingleNode("PERSONGUID")
    If xmlNode Is Nothing Then
        Set xmlElement = vXMLRequest.ownerDocument.createElement("PERSONGUID")
        vXMLRequest.appendChild xmlElement
    Else
        Set xmlElement = xmlNode
    End If
        
    xmlElement.Text = strPersonGuid
        
    Set objIClassDef = New omCustClassDef
    objIDOAssist.Update vXMLRequest, objIClassDef.LoadDependantData

    '
    ' Delete Person
    '
    If botPersonOperation = booDelete Then
        ' Person should be deleted only once the thirdparty record has been updated
        SavePerson vXMLRequest, strPersonGuid, vbotOperationType:=botPersonOperation
    End If

    Set objIDOAssist = Nothing
    Set objIClassDef = Nothing
    Set objIPersonClassDef = Nothing
    Set objErrAssist = Nothing
    Set objXmlAssist = Nothing
    Set xmlNode = Nothing
    Set xmlElement = Nothing

    If Not objContext Is Nothing Then
        objContext.SetComplete
        Set objContext = Nothing
    End If

    Exit Function

ICustomerDO_UpdateDependantVbErr:

    Set objIDOAssist = Nothing
    Set objXmlAssist = Nothing
    Set xmlNode = Nothing
    Set objIClassDef = Nothing
    Set objIPersonClassDef = Nothing
    Set xmlElement = Nothing
    
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
        Set objContext = Nothing
    End If

    Set objErrAssist = Nothing
    '   re-raise error for business object to interpret as appropriate
    Err.Raise Err.Number, Err.Source, Err.Description

End Function

Public Sub ICustomerDO_UpdatePersonalDetails(ByVal vXMLRequest As IXMLDOMElement)
' header ----------------------------------------------------------------------------------
' description:
'   Update a single instance of the persistant data associated with this
'   data object
' pass:
'   vstrXMLRequest  xml Request data stream containing data to be persisted
' Raise Errors:
'------------------------------------------------------------------------------------------
On Error GoTo ICustomerDO_UpdatePersonalDetailsVbErr

    Const strFunctionName As String = "ICustomerDO_UpdatePersonalDetails"

    Dim objErrAssist As New ErrAssist
    Dim objXmlAssist As New XMLAssist
    Dim objComboDO As New ComboDO

    Dim xmlNodeList As IXMLDOMNodeList
    Dim xmlNode As IXMLDOMNode
    Dim xmlCustomerVersionElem As IXMLDOMElement

    Dim strTitle      As String, _
        strSurname    As String, _
        strTitleId    As String

    Dim objIClassDef As IomCustClassDef
    
    Dim strXML As String

    Dim objContext As ObjectContext
    Set objContext = GetObjectContext()

    Dim objIDOAssist As IDOAssist

    If objContext Is Nothing Then
        Set objIDOAssist = New DOAssist
        Set objComboDO = New ComboDO
    Else
        Set objIDOAssist = objContext.CreateInstance(gstrBASE_COMPONENT & ".DOAssist")
        Set objComboDO = objContext.CreateInstance(gstrBASE_COMPONENT & ".ComboDO")
    End If

'    '
'    ' Default text for salutation if it is null
'    '
'    Set xmlNodeList = vxmlRequest.getElementsByTagName("CORRESPONDENCESALUTATION")
'    If xmlNodeList.length > 0 Then
'        Set xmlNode = xmlNodeList.Item(0)
'
'        ' There should actually only ever be one node, at most
'        If xmlNode.Text = "" Then
'            ' Salutation is null, so substitute in a default value
'            ' Get surname
'            strSurname = objXmlAssist.GetTagValue(vxmlRequest, "SURNAME")
'
'            ' Get title
'            strTitle = objXmlAssist.GetTagValue(vxmlRequest, "TITLEOTHER")
'            If strTitle = "" Then
'                ' No custom title is specified so get the text corresponding to the specified Title ID
'                strTitleId = objXmlAssist.GetTagValue(vxmlRequest, "TITLE")
'                strTitle = objComboDO.GetComboText("Title", strTitleId)
'            End If
'
'            xmlNode.Text = strTitle & " " & strSurname
'        End If
'    End If

    '
    ' Update Customer
    '
    Set objIClassDef = New omCustClassDef
    
    ' Update Customer Version data
    Set xmlCustomerVersionElem = vXMLRequest.selectSingleNode("CUSTOMERVERSION")
    AddSalutationToRequest xmlCustomerVersionElem
    objIDOAssist.Update xmlCustomerVersionElem, objIClassDef.LoadCustomerVersionData


    ' Update Customer Data
    objIDOAssist.Update vXMLRequest, objIClassDef.LoadCustomerData


    Set objErrAssist = Nothing
    Set objIDOAssist = Nothing
    Set objXmlAssist = Nothing
    Set objIClassDef = Nothing
    Set xmlCustomerVersionElem = Nothing
    Set xmlNodeList = Nothing
    Set xmlNode = Nothing
    Set objComboDO = Nothing

    If Not objContext Is Nothing Then
        objContext.SetComplete
        Set objContext = Nothing
    End If


    Exit Sub

ICustomerDO_UpdatePersonalDetailsVbErr:

    Set objIDOAssist = Nothing
    Set objXmlAssist = Nothing
    Set xmlCustomerVersionElem = Nothing
    Set xmlNodeList = Nothing
    Set xmlNode = Nothing
    Set objComboDO = Nothing

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
        Set objContext = Nothing
    End If

    Set objErrAssist = Nothing
    '   re-raise error for business object to interpret as appropriate
    Err.Raise Err.Number, Err.Source, Err.Description

End Sub

Private Sub SavePerson(ByVal vXMLRequest As IXMLDOMElement, _
                        ByRef rstrPersonGUID As String, _
                        Optional ByVal vbotOperationType As Variant)
' header ----------------------------------------------------------------------------------
' description:
'   Update a single instance of the persistant data associated with this
'   data object
' pass:
'   vXMLRequest  xml Request node containing data to be persisted
' Raise Errors:
'------------------------------------------------------------------------------------------
On Error GoTo SavePersonVbErr

    Const strFunctionName As String = "SavePerson"

    Dim objErrAssist As New ErrAssist

    Dim objContext As ObjectContext
    Set objContext = GetObjectContext()

    Dim objXmlAssist As New XMLAssist
    Dim xmlTemp As FreeThreadedDOMDocument40

    Dim botOperationType As BOOPERATIONTYPE
    Dim objIPersonDO As IPersonDO
    Dim xmlReturnElem As IXMLDOMElement
    Dim xmlPersonOutElem As IXMLDOMElement ' Person node in Output
    Dim xmlPersonInElem As IXMLDOMElement ' Person node in Input

    If objContext Is Nothing Then
        Set objIPersonDO = New PersonDO
    Else
        Set objIPersonDO = objContext.CreateInstance(cstrPERSON_COMPONENT & ".PersonDO")
    End If
    '
    ' Initialise
    '
    rstrPersonGUID = ""
    botOperationType = booNone
    If Not IsMissing(vbotOperationType) Then botOperationType = vbotOperationType
    '
    ' Main
    '
    If vXMLRequest.nodeName = "PERSON" Then
        Set xmlPersonInElem = vXMLRequest
    Else
        Set xmlPersonInElem = vXMLRequest.selectSingleNode(".//PERSON")
    End If
    If Not objXmlAssist.GetTagValue(vXMLRequest, "PERSON") = "" Then
        Set xmlReturnElem = objIPersonDO.Save(xmlPersonInElem, rbotOperationType:=botOperationType)

        ' Extract the Person GUID from the appropriate source
        Select Case botOperationType
            Case booCreate
                rstrPersonGUID = objXmlAssist.GetTagValue(xmlReturnElem, "PERSONGUID")
            Case booDelete, booNone
                rstrPersonGUID = ""
            Case booUpdate
                Set xmlPersonOutElem = vXMLRequest.selectSingleNode(".//PERSON")
                rstrPersonGUID = objXmlAssist.GetTagValue(xmlPersonOutElem, "PERSONGUID")
            Case Else
                ' N/A
        End Select
    End If

    Set objErrAssist = Nothing

    Set objXmlAssist = Nothing
    Set xmlPersonOutElem = Nothing
    Set xmlReturnElem = Nothing
    Set xmlTemp = Nothing

    Set objIPersonDO = Nothing

    If Not objContext Is Nothing Then
        objContext.SetComplete
    End If


    Exit Sub

SavePersonVbErr:

    Set objXmlAssist = Nothing
    Set xmlReturnElem = Nothing
    Set xmlPersonOutElem = Nothing

    Set objIPersonDO = Nothing

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

Private Sub ICustomerDO_DeleteAllCustomerTelephone(ByVal vXMLRequest As IXMLDOMElement)
' header ----------------------------------------------------------------------------------
' description:
'   Delete all instances of the persistant data associated with this
'   data object that match the key values specified
' pass:
'   vstrXMLRequest  xml Request data stream containing data to which identifies
'                   the instance(s) to be deleted
' Raise Errors:
'------------------------------------------------------------------------------------------
On Error GoTo ICustomerDO_DeleteAllCustomerTelephoneVbErr

    Const strFunctionName As String = "ICustomerDO_DeleteAllCustomerTelephone"

    Dim objErrAssist As New ErrAssist

    Dim objIDOAssist As IDOAssist
    Dim objIClassDef As IomCustClassDef

    Dim objContext As ObjectContext
    Set objContext = GetObjectContext()

    If objContext Is Nothing Then
        Set objIDOAssist = New DOAssist
    Else
        Set objIDOAssist = objContext.CreateInstance(gstrBASE_COMPONENT & ".DOAssist")
    End If
    
    Set objIClassDef = New omCustClassDef
    objIDOAssist.DeleteAll vXMLRequest, objIClassDef.LoadCustomerTelephoneData

    Set objIDOAssist = Nothing
    Set objIClassDef = Nothing
    Set objErrAssist = Nothing

    If Not objContext Is Nothing Then
        objContext.SetComplete
        Set objContext = Nothing
    End If

    Exit Sub

ICustomerDO_DeleteAllCustomerTelephoneVbErr:

    Set objIDOAssist = Nothing
    Set objIClassDef = Nothing
    
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
        Set objContext = Nothing
        Set objContext = Nothing
    End If

    Set objErrAssist = Nothing

    '   re-raise error for business object to interpret as appropriate
    Err.Raise Err.Number, Err.Source, Err.Description

End Sub

Private Function ICustomerDO_FindDependantList(ByVal vXMLRequest As IXMLDOMElement) As IXMLDOMElement
' header ----------------------------------------------------------------------------------
' description:
'   Get the data for all instances of the persistant data associated with
'   this data object for the values supplied
' pass:
'   vXMLRequest  xml Request node containing data to which identifies
'                   the instance(s) of the persistant data to be retrieved
' return:
'   FindList        XML node representing the  data retrieved
' Raise Errors:
'   if record not found, raise error
'------------------------------------------------------------------------------------------
On Error GoTo ICustomerDO_FindDependantListVbErr

    Const strFunctionName As String = "ICustomerDO_FindDependantList"

    Dim objErrAssist As New ErrAssist
    Dim objXmlAssist As New XMLAssist

    Dim strXML As String

    Dim objIDOAssist As IDOAssist
    Dim objIClassDef As IomCustClassDef

    Dim objContext As ObjectContext
    Set objContext = GetObjectContext()
    Dim xmlDoc As FreeThreadedDOMDocument40
    Dim xmlElement As IXMLDOMElement
    Dim xmlNode As IXMLDOMNode
    
    If objContext Is Nothing Then
        Set objIDOAssist = New DOAssist
    Else
        Set objIDOAssist = objContext.CreateInstance(gstrBASE_COMPONENT & ".DOAssist")
    End If

    strXML = vXMLRequest.xml
    strXML = Replace(strXML, "<DEPENDANT>", "<DEPENDANTDETAILS>")
    strXML = Replace(strXML, "</DEPENDANT>", "</DEPENDANTDETAILS>")
    
    Set xmlDoc = objXmlAssist.load(strXML, TypeName(Me), strFunctionName)
    Set objIClassDef = New omCustClassDef
   
    Set xmlElement = objIDOAssist.FindListMultiple(xmlDoc.documentElement, _
                                            objIClassDef.LoadDependantDetailsData, _
                                            vstrItemName:="DEPENDANTDETAILS")
       
    ' Format the out put in a hierarchial way, i.e, 'PERSON' node comes beneath 'DEPENDANT'
    For Each xmlNode In xmlElement.selectNodes("DEPENDANTDETAILS/DEPENDANT")
        objXmlAssist.PromoteNode xmlElement, xmlNode
    Next xmlNode
       
    Set ICustomerDO_FindDependantList = xmlElement.cloneNode(True)

    Set objErrAssist = Nothing
    Set objIDOAssist = Nothing
    Set objXmlAssist = Nothing
    Set xmlDoc = Nothing
    Set xmlElement = Nothing
    Set xmlNode = Nothing

    If Not objContext Is Nothing Then
        objContext.SetComplete
        Set objContext = Nothing
    End If

    Exit Function

ICustomerDO_FindDependantListVbErr:

    Set objIDOAssist = Nothing
    Set objXmlAssist = Nothing
    Set xmlDoc = Nothing
    Set xmlElement = Nothing
    Set xmlNode = Nothing
    
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
        Set objContext = Nothing
    End If

    Set objErrAssist = Nothing

    '   re-raise error for business object to interpret as appropriate
    Err.Raise Err.Number, Err.Source, Err.Description

End Function

Private Sub ICustomerDO_UpdateCustomerTelephone(ByVal vXMLRequest As IXMLDOMElement)
' header ----------------------------------------------------------------------------------
' description:
'   Update a single instance of the persistant data associated with this
'   data object
' pass:
'   vstrXMLRequest  xml Request data stream containing data to be persisted
' Raise Errors:
'------------------------------------------------------------------------------------------
On Error GoTo ICustomerDO_UpdateCustomerTelephoneVbErr

    Const strFunctionName As String = "ICustomerDO_UpdateCustomerTelephone"

    Dim objErrAssist As New ErrAssist

    Dim objContext As ObjectContext
    Set objContext = GetObjectContext()

    Dim objIDOAssist As IDOAssist
    Dim objIClassDef As IomCustClassDef

    If objContext Is Nothing Then
      Set objIDOAssist = New DOAssist
    Else
        Set objIDOAssist = objContext.CreateInstance(gstrBASE_COMPONENT & ".DOAssist")
    End If
    
    Set objIClassDef = New omCustClassDef
    objIDOAssist.Update vXMLRequest, objIClassDef.LoadCustomerTelephoneData

    Set objIDOAssist = Nothing
    Set objErrAssist = Nothing

    If Not objContext Is Nothing Then
        objContext.SetComplete
        Set objContext = Nothing
    End If

    Exit Sub

ICustomerDO_UpdateCustomerTelephoneVbErr:

    Set objIDOAssist = Nothing
    Set objIClassDef = Nothing

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
        Set objContext = Nothing
    End If

    Set objErrAssist = Nothing
    '   re-raise error for business object to interpret as appropriate
    Err.Raise Err.Number, Err.Source, Err.Description
    
End Sub

Private Sub ICustomerDO_DeleteCustomerTelephone(ByVal vXMLRequest As IXMLDOMElement)
' header ----------------------------------------------------------------------------------
' description:
'   Delete a single instance of the persistant data associated with this
'   data object
' pass:
'   vXMLRequest  xml Request data stream containing data to which identifies
'                the instance to be deleted
' Raise Errors:
'       ApplicationRecordNotFound
'       ApplicationInvalidKeyString
'       parser errors
'------------------------------------------------------------------------------------------
On Error GoTo ICustomerDO_DeleteCustomerTelephoneVbErr

    Const strFunctionName As String = "ICustomerDO_DeleteCustomerTelephone"

    Dim objErrAssist As New ErrAssist

    Dim objContext As ObjectContext
    Set objContext = GetObjectContext()

    Dim objIDOAssist As IDOAssist
    Dim objIClassDef As IomCustClassDef

    If objContext Is Nothing Then
        Set objIDOAssist = New DOAssist
    Else
        Set objIDOAssist = objContext.CreateInstance(gstrBASE_COMPONENT & ".DOAssist")
    End If
    
    Set objIClassDef = New omCustClassDef
    objIDOAssist.Delete vXMLRequest, objIClassDef.LoadCustomerTelephoneData

    Set objIDOAssist = Nothing
    Set objIClassDef = Nothing
    Set objErrAssist = Nothing

    If Not objContext Is Nothing Then
        objContext.SetComplete
        Set objContext = Nothing
    End If

    Exit Sub

ICustomerDO_DeleteCustomerTelephoneVbErr:

    Set objIDOAssist = Nothing
    Set objIClassDef = Nothing
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
        Set objContext = Nothing
    End If

    Set objErrAssist = Nothing
    '   re-raise error for business object to interpret as appropriate
    Err.Raise Err.Number, Err.Source, Err.Description

End Sub

Private Sub ICustomerDO_UpdateWrapUpDetails(ByVal vXMLRequest As IXMLDOMElement)
' header ----------------------------------------------------------------------------------
' description:
'   Update a single instance of the persistant data associated with this
'   data object
' pass:
'   vstrXMLRequest  xml Request data stream containing data to be persisted
' Raise Errors:
'------------------------------------------------------------------------------------------
On Error GoTo ICustomerDO_UpdateWrapUpDetailsVbErr

    Const strFunctionName As String = "ICustomerDO_UpdateWrapUpDetails"

    Dim objErrAssist As New ErrAssist

    Dim objContext As ObjectContext
    Set objContext = GetObjectContext()

    Dim objIDOAssist As IDOAssist
    Dim objIClassDef As IomCustClassDef

    If objContext Is Nothing Then
      Set objIDOAssist = New DOAssist
    Else
        Set objIDOAssist = objContext.CreateInstance(gstrBASE_COMPONENT & ".DOAssist")
    End If
    
    Set objIClassDef = New omCustClassDef
    objIDOAssist.Update vXMLRequest, objIClassDef.LoadWrapUpDetailsData

    Set objIDOAssist = Nothing
    Set objErrAssist = Nothing

    If Not objContext Is Nothing Then
        objContext.SetComplete
        Set objContext = Nothing
    End If

    Exit Sub

ICustomerDO_UpdateWrapUpDetailsVbErr:

    Set objIDOAssist = Nothing
    Set objIClassDef = Nothing

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
        Set objContext = Nothing
    End If

    Set objErrAssist = Nothing
    '   re-raise error for business object to interpret as appropriate
    Err.Raise Err.Number, Err.Source, Err.Description
    
End Sub

Private Sub ICustomerDO_DeleteWrapUpDetails(ByVal vXMLRequest As IXMLDOMElement)
' header ----------------------------------------------------------------------------------
' description:
'   Delete a single instance of the persistant data associated with this
'   data object
' pass:
'   vXMLRequest  xml Request data stream containing data to which identifies
'                the instance to be deleted
' Raise Errors:
'       ApplicationRecordNotFound
'       ApplicationInvalidKeyString
'       parser errors
'------------------------------------------------------------------------------------------
On Error GoTo ICustomerDO_DeleteWrapUpDetailsVbErr

    Const strFunctionName As String = "ICustomerDO_DeleteWrapUpDetails"

    Dim objErrAssist As New ErrAssist

    Dim objContext As ObjectContext
    Set objContext = GetObjectContext()

    Dim objIDOAssist As IDOAssist
    Dim objIClassDef As IomCustClassDef

    If objContext Is Nothing Then
        Set objIDOAssist = New DOAssist
    Else
        Set objIDOAssist = objContext.CreateInstance(gstrBASE_COMPONENT & ".DOAssist")
    End If
    
    Set objIClassDef = New omCustClassDef
    objIDOAssist.Delete vXMLRequest, objIClassDef.LoadWrapUpDetailsData

    Set objIDOAssist = Nothing
    Set objIClassDef = Nothing
    Set objErrAssist = Nothing

    If Not objContext Is Nothing Then
        objContext.SetComplete
        Set objContext = Nothing
    End If

    Exit Sub

ICustomerDO_DeleteWrapUpDetailsVbErr:

    Set objIDOAssist = Nothing
    Set objIClassDef = Nothing
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
        Set objContext = Nothing
    End If

    Set objErrAssist = Nothing
    '   re-raise error for business object to interpret as appropriate
    Err.Raise Err.Number, Err.Source, Err.Description

End Sub


Private Sub ICustomerDO_FindCustomer(ByVal vobjXmlRequest As IXMLDOMNode, _
                                     ByVal vobjXmlResponse As IXMLDOMNode)

On Error GoTo FindCustomerVbErr

    Const strFunctionName As String = "ICustomerDO_FindCustomer"
    
    Dim objSQLAssist As New SQLAssist
    Dim objADOAssist As New ADOAssist
    Dim objErrAssist As New ErrAssist
    
    Dim adoConnection As New ADODB.Connection
    Dim adoCommand As New ADODB.Command
    Dim adoParameter As ADODB.Parameter
    Dim rstThisRecSet As New ADODB.Recordset
    
    Dim xmlNode As IXMLDOMNode
    Dim xmlElem As IXMLDOMElement
    
    Dim strSQL As String
    Dim strSurname As String, strForename As String, strDoB As String, strPostcode As String
    Dim strSecondForeName As String, strOtherForeNames As String
    Dim blnIsWildCard As Boolean
    Dim strPageNum As String, strPageSize As String
    Dim intStartRow As Integer, intPageNum As Integer, intPageSize As Integer
    Dim intRows As Integer
    Dim strHouseName As String
    Dim strHouseNumber As String
    Dim strChannelId As String
    Dim strFlatNumber As String
    Dim strOtherSysCustNo As String
    Dim strPhoneNo As String
        
    intStartRow = 1
    intPageSize = cintFINDCUSTOMERPAGE

    Set xmlNode = vobjXmlRequest.selectSingleNode("//SEARCH")
    
    If Not xmlNode Is Nothing Then
        
        If Not xmlNode.Attributes.getNamedItem("PAGESIZE") Is Nothing Then
            
            strPageSize = xmlNode.Attributes.getNamedItem("PAGESIZE").Text
            
            If Len(strPageSize) > 0 Then
                intPageSize = CInt(strPageSize)
            End If
        
        End If
        
        If intPageSize < 1 Then
            intPageSize = cintFINDCUSTOMERPAGE
        End If
        
        If Not xmlNode.Attributes.getNamedItem("PAGENUMBER") Is Nothing Then
            
            strPageNum = xmlNode.Attributes.getNamedItem("PAGENUMBER").Text
            
            If Len(strPageNum) > 0 Then
                
                intPageNum = CInt(strPageNum)
                
                If intPageNum < 1 Then
                    intPageNum = 1
                End If
                
                intStartRow = ((intPageNum - 1) * intPageSize) + 1
            
            End If
        
        End If
    
    End If
    
    ' SR : 02/05/200 - AQR/SYS0227 : Do not check for Aposterphes in the values
    '                  assigned to parameters of Command Object
    Set xmlNode = vobjXmlRequest.selectSingleNode("//SURNAME")
    If Not xmlNode Is Nothing Then
        strSurname = xmlNode.Text
        strSurname = objSQLAssist.FormatWildcardedString(strSurname, blnIsWildCard, True)
    End If
    
    Set xmlNode = vobjXmlRequest.selectSingleNode("//FIRSTFORENAME")
    If Not xmlNode Is Nothing Then
        strForename = xmlNode.Text
        strForename = objSQLAssist.FormatWildcardedString(strForename, blnIsWildCard, True)
    End If
    
    Set xmlNode = vobjXmlRequest.selectSingleNode("//SECONDFORENAME")
    If Not xmlNode Is Nothing Then
        strSecondForeName = xmlNode.Text
        strSecondForeName = objSQLAssist.FormatWildcardedString(strSecondForeName, blnIsWildCard, True)
    End If
    
    Set xmlNode = vobjXmlRequest.selectSingleNode("//OTHERFORENAMES")
    If Not xmlNode Is Nothing Then
        strOtherForeNames = xmlNode.Text
        strOtherForeNames = objSQLAssist.FormatWildcardedString(strOtherForeNames, blnIsWildCard, True)
    End If
    
    Set xmlNode = vobjXmlRequest.selectSingleNode("//DATEOFBIRTH")
    If Not xmlNode Is Nothing Then
        strDoB = xmlNode.Text
    End If
    
    Set xmlNode = vobjXmlRequest.selectSingleNode("//POSTCODE")
    If Not xmlNode Is Nothing Then
        strPostcode = xmlNode.Text
        strPostcode = objSQLAssist.FormatWildcardedString(strPostcode, blnIsWildCard, True)
    End If
    
    ' PSC 09/01/01 SYS1750 - Start
    strHouseName = m_objXmlAssist.GetNodeValue(vobjXmlRequest, "//BUILDINGORHOUSENAME", False, False)
    strHouseNumber = m_objXmlAssist.GetNodeValue(vobjXmlRequest, "//BUILDINGORHOUSENUMBER", False, False)
    strChannelId = m_objXmlAssist.GetAttributeFromNode(vobjXmlRequest, "CHANNELID")
    strFlatNumber = m_objXmlAssist.GetNodeValue(vobjXmlRequest, "//FLATNUMBER", False, False)
    strOtherSysCustNo = m_objXmlAssist.GetNodeValue(vobjXmlRequest, "//OTHERSYSTEMCUSTOMERNUMBER", False, False)
    strPhoneNo = m_objXmlAssist.GetNodeValue(vobjXmlRequest, "//CUSTOMERTELEPHONENUMBER", False, False)
    ' PSC 09/01/01 SYS1750 - End
    
    Set xmlNode = Nothing
    
    adoConnection.ConnectionString = objADOAssist.GetConnStr
    adoConnection.CursorLocation = adUseClient
    adoConnection.Open
    
    Set adoCommand.ActiveConnection = adoConnection
    adoCommand.CommandType = adCmdText
    
#If GENERIC_SQL Then
    ' AS 30/05/01 CC012 - Start
    Dim eDbEngineType As DBENGINETYPE
    eDbEngineType = objADOAssist.GetDBEngine
    If eDbEngineType = SQLServer Then
        strSQL = "{call usp_CustomerRegNoRowsReturned(?,?,?,?,?,?,?,?,?,?,?,?,?)}"
    ElseIf eDbEngineType = Oracle Then
        ' PSC 09/01/01 SYS1750 - Start
        strSQL = "{call sp_CustomerReg.NOROWSRESTURNED(?,?,?,?,?,?,?,?,?,?,?,?,"
        strSQL = strSQL & "{resultset 1,"
        strSQL = strSQL & "NOROWS"
        strSQL = strSQL & "})}"
        ' PSC 09/01/01 SYS1750 - End
    End If
    ' AS 30/05/01 CC012 - End
#Else
    ' PSC 09/01/01 SYS1750 - Start
    strSQL = "{call sp_CustomerReg.NOROWSRESTURNED(?,?,?,?,?,?,?,?,?,?,?,?,"
    strSQL = strSQL & "{resultset 1,"
    strSQL = strSQL & "NOROWS"
    strSQL = strSQL & "})}"
    ' PSC 09/01/01 SYS1750 - End
#End If

    adoCommand.CommandText = strSQL

' AS 30/05/01 CC012 - Start
#If GENERIC_SQL Then

    ' AS Initialise parameter values to Null.
    Set adoParameter = adoCommand.CreateParameter(, adBSTR, adParamInput, Len(strSurname), Null)
    adoCommand.Parameters.Append adoParameter

    Set adoParameter = adoCommand.CreateParameter(, adBSTR, adParamInput, Len(strForename), Null)
    adoCommand.Parameters.Append adoParameter

    Set adoParameter = adoCommand.CreateParameter(, adBSTR, adParamInput, Len(strSecondForeName), Null)
    adoCommand.Parameters.Append adoParameter

    Set adoParameter = adoCommand.CreateParameter(, adBSTR, adParamInput, Len(strOtherForeNames), Null)
    adoCommand.Parameters.Append adoParameter

    Set adoParameter = adoCommand.CreateParameter(, adBSTR, adParamInput, Len(strDoB), Null)
    adoCommand.Parameters.Append adoParameter

    Set adoParameter = adoCommand.CreateParameter(, adBSTR, adParamInput, Len(strPostcode), Null)
    adoCommand.Parameters.Append adoParameter
    
    ' PSC 09/01/01 SYS1750 - Start
    Set adoParameter = adoCommand.CreateParameter(, adBSTR, adParamInput, Len(strHouseName), Null)
    adoCommand.Parameters.Append adoParameter
    
    Set adoParameter = adoCommand.CreateParameter(, adBSTR, adParamInput, Len(strHouseNumber), Null)
    adoCommand.Parameters.Append adoParameter

    Set adoParameter = adoCommand.CreateParameter(, adBSTR, adParamInput, Len(strFlatNumber), Null)
    adoCommand.Parameters.Append adoParameter

    Set adoParameter = adoCommand.CreateParameter(, adBSTR, adParamInput, Len(strPhoneNo), Null)
    adoCommand.Parameters.Append adoParameter

    Set adoParameter = adoCommand.CreateParameter(, adBSTR, adParamInput, Len(strChannelId), Null)
    adoCommand.Parameters.Append adoParameter

    Set adoParameter = adoCommand.CreateParameter(, adBSTR, adParamInput, Len(strOtherSysCustNo), Null)
    adoCommand.Parameters.Append adoParameter
    ' PSC 09/01/01 SYS1750 - End
    
    ' For SQL Server, the number of rows is returned as an output parameter.
    ' For Oracle, the number of rows is returned in the resultset (see above).
    If eDbEngineType = SQLServer Then
        Set adoParameter = adoCommand.CreateParameter(, adInteger, adParamOutput)
        adoCommand.Parameters.Append adoParameter
    End If

    ' AS Now override initial Null parameter value with string - but only if string is not empty.
    ' If the string is empty then it will stay Null.
    If Not strSurname = "" Then
        adoCommand(0) = strSurname
    End If
    If Not strForename = "" Then
        adoCommand(1) = strForename
    End If
    If Not strSecondForeName = "" Then
        adoCommand(2) = strSecondForeName
    End If
    If Not strOtherForeNames = "" Then
        adoCommand(3) = strOtherForeNames
    End If
    If Not strDoB = "" Then
        adoCommand(4) = strDoB
    End If
    If Not strPostcode = "" Then
        adoCommand(5) = strPostcode
    End If
    
    ' PSC 09/01/01 SYS1750 - Start
    If Not strHouseName = "" Then
        adoCommand(6) = strHouseName
    End If
    If Not strHouseNumber = "" Then
        adoCommand(7) = strHouseNumber
    End If
    If Not strFlatNumber = "" Then
        adoCommand(8) = strFlatNumber
    End If
    If Not strPhoneNo = "" Then
        adoCommand(9) = strPhoneNo
    End If
    If Not strChannelId = "" Then
        adoCommand(10) = strChannelId
    End If
    If Not strOtherSysCustNo = "" Then
        adoCommand(11) = strOtherSysCustNo
    End If
    ' PSC 09/01/01 SYS1750 - End
    
#Else

    Set adoParameter = adoCommand.CreateParameter(, adBSTR, adParamInput, Len(strSurname), strSurname)
    adoCommand.Parameters.Append adoParameter

    Set adoParameter = adoCommand.CreateParameter(, adBSTR, adParamInput, Len(strForename), strForename)
    adoCommand.Parameters.Append adoParameter

    Set adoParameter = adoCommand.CreateParameter(, adBSTR, adParamInput, Len(strSecondForeName), strSecondForeName)
    adoCommand.Parameters.Append adoParameter

    Set adoParameter = adoCommand.CreateParameter(, adBSTR, adParamInput, Len(strOtherForeNames), strOtherForeNames)
    adoCommand.Parameters.Append adoParameter

    Set adoParameter = adoCommand.CreateParameter(, adBSTR, adParamInput, Len(strDoB), strDoB)
    adoCommand.Parameters.Append adoParameter

    Set adoParameter = adoCommand.CreateParameter(, adBSTR, adParamInput, Len(strPostcode), strPostcode)
    adoCommand.Parameters.Append adoParameter
    
    ' PSC 09/01/01 SYS1750 - Start
    Set adoParameter = adoCommand.CreateParameter(, adBSTR, adParamInput, Len(strHouseName), strHouseName)
    adoCommand.Parameters.Append adoParameter
    
    Set adoParameter = adoCommand.CreateParameter(, adBSTR, adParamInput, Len(strHouseNumber), strHouseNumber)
    adoCommand.Parameters.Append adoParameter

    Set adoParameter = adoCommand.CreateParameter(, adBSTR, adParamInput, Len(strFlatNumber), strFlatNumber)
    adoCommand.Parameters.Append adoParameter

    Set adoParameter = adoCommand.CreateParameter(, adBSTR, adParamInput, Len(strPhoneNo), strPhoneNo)
    adoCommand.Parameters.Append adoParameter

    Set adoParameter = adoCommand.CreateParameter(, adBSTR, adParamInput, Len(strChannelId), strChannelId)
    adoCommand.Parameters.Append adoParameter

    Set adoParameter = adoCommand.CreateParameter(, adBSTR, adParamInput, Len(strOtherSysCustNo), strOtherSysCustNo)
    adoCommand.Parameters.Append adoParameter

#End If
' AS 30/05/01 CC012 - End
                
' AS 30/05/01 CC012 - Start
#If GENERIC_SQL Then
    If eDbEngineType = SQLServer Then
        ' AS 04/07/01 CC012 SQL Server usp_CustomerRegNoRowsReturned stored procedure no longer
        ' called as usp_CustomerRegFindCustomerDetails returns the total number of matching rows.
        'Set rstThisRecSet = adoCommand.Execute
        'intRows = adoCommand(12)
        'If intRows = 0 Then
        '    objErrAssist.ThrowError TypeName(Me), strFunctionName, oeRecordNotFound
        'End If
    ElseIf eDbEngineType = Oracle Then
        Set rstThisRecSet = adoCommand.Execute
        If rstThisRecSet.EOF Then  'This condition will never be true, but just in case
            ' raise application error to be interpreted by calling object
            objErrAssist.ThrowError TypeName(Me), strFunctionName, oeRecordNotFound
        Else
            intRows = rstThisRecSet.Fields.Item("NOROWS").Value
            If intRows = 0 Then
                objErrAssist.ThrowError TypeName(Me), strFunctionName, oeRecordNotFound
            End If
        End If
    End If
#Else
    Set rstThisRecSet = adoCommand.Execute
    If rstThisRecSet.EOF Then  'This condition will never be true, but just in case
        ' raise application error to be interpreted by calling object
        objErrAssist.ThrowError TypeName(Me), strFunctionName, oeRecordNotFound
    Else
        intRows = rstThisRecSet.Fields.Item("NOROWS").Value
        If intRows = 0 Then
            objErrAssist.ThrowError TypeName(Me), strFunctionName, oeRecordNotFound
        End If
    End If
#End If
' AS 30/05/01 CC012 - End
    
    If rstThisRecSet.State = adStateOpen Then
        rstThisRecSet.Close
    End If
    
' AS 30/05/01 CC012 - Start
#If GENERIC_SQL Then
    If eDbEngineType = SQLServer Then
        strSQL = "{call usp_CustomerRegFindCustomerDetails(?,?,?,?,?,?,?,?,?,?,?,?,?,?,?)}"
    ElseIf eDbEngineType = Oracle Then
        ' PSC 09/01/01 SYS1750 - Start
        strSQL = "{call sp_CustomerReg.FindCustomerDetails(?,?,?,?,?,?,?,?,?,?,?,?,?,?,"
        strSQL = strSQL & "{resultset " & CStr(intPageSize + 1) & ","
        strSQL = strSQL & "CUSTOMERNUMBER,"
        strSQL = strSQL & "CUSTOMERVERSIONNUMBER,"
        strSQL = strSQL & "CORRESPONDENCESALUTATION,"
        strSQL = strSQL & "DATEOFBIRTH,"
        strSQL = strSQL & "FIRSTFORENAME,"
        strSQL = strSQL & "OTHERFORENAMES,"
        strSQL = strSQL & "SECONDFORENAME,"
        strSQL = strSQL & "SURNAME,"
        strSQL = strSQL & "OTHERSYSTEMCUSTOMERNUMBER,"
        strSQL = strSQL & "OTHERSYSTEMTYPE,"
        strSQL = strSQL & "CHANNELID,"
        strSQL = strSQL & "BUILDINGORHOUSENAME,"
        strSQL = strSQL & "BUILDINGORHOUSENUMBER,"
        strSQL = strSQL & "FLATNUMBER,"
        strSQL = strSQL & "STREET,"
        strSQL = strSQL & "DISTRICT,"
        strSQL = strSQL & "TOWN,"
        strSQL = strSQL & "COUNTY,"
        strSQL = strSQL & "COUNTRY,"
        strSQL = strSQL & "POSTCODE,"
        strSQL = strSQL & "ADDRESSTYPE,"
        strSQL = strSQL & "MOTHERSMAIDENNAME"
        strSQL = strSQL & "})}"
        ' PSC 09/01/01 SYS1750 - End
    End If
#Else
    ' PSC 09/01/01 SYS1750 - Start
    strSQL = "{call sp_CustomerReg.FindCustomerDetails(?,?,?,?,?,?,?,?,?,?,?,?,?,?,"
    strSQL = strSQL & "{resultset " & CStr(intPageSize + 1) & ","
    strSQL = strSQL & "CUSTOMERNUMBER,"
    strSQL = strSQL & "CUSTOMERVERSIONNUMBER,"
    strSQL = strSQL & "CORRESPONDENCESALUTATION,"
    strSQL = strSQL & "DATEOFBIRTH,"
    strSQL = strSQL & "FIRSTFORENAME,"
    strSQL = strSQL & "OTHERFORENAMES,"
    strSQL = strSQL & "SECONDFORENAME,"
    strSQL = strSQL & "SURNAME,"
    strSQL = strSQL & "OTHERSYSTEMCUSTOMERNUMBER,"
    strSQL = strSQL & "OTHERSYSTEMTYPE,"
    strSQL = strSQL & "CHANNELID,"
    strSQL = strSQL & "BUILDINGORHOUSENAME,"
    strSQL = strSQL & "BUILDINGORHOUSENUMBER,"
    strSQL = strSQL & "FLATNUMBER,"
    strSQL = strSQL & "STREET,"
    strSQL = strSQL & "DISTRICT,"
    strSQL = strSQL & "TOWN,"
    strSQL = strSQL & "COUNTY,"
    strSQL = strSQL & "COUNTRY,"
    strSQL = strSQL & "POSTCODE,"
    strSQL = strSQL & "ADDRESSTYPE,"
    strSQL = strSQL & "MOTHERSMAIDENNAME"
    strSQL = strSQL & "})}"
    ' PSC 09/01/01 SYS1750 - End
#End If
' AS 30/05/01 CC012 - End

    Set adoCommand = Nothing
    Set adoCommand = New ADODB.Command
    Set adoCommand.ActiveConnection = adoConnection
    
    adoCommand.CommandText = strSQL

' AS 30/05/01 CC012 - Start
#If GENERIC_SQL Then
    Set adoParameter = adoCommand.CreateParameter(, adInteger, adParamInput)
    adoCommand.Parameters.Append adoParameter

    Set adoParameter = adoCommand.CreateParameter(, adInteger, adParamInput)
    adoCommand.Parameters.Append adoParameter

    Set adoParameter = adoCommand.CreateParameter(, adBSTR, adParamInput, Len(strSurname), Null)
    adoCommand.Parameters.Append adoParameter

    Set adoParameter = adoCommand.CreateParameter(, adBSTR, adParamInput, Len(strForename), Null)
    adoCommand.Parameters.Append adoParameter

    Set adoParameter = adoCommand.CreateParameter(, adBSTR, adParamInput, Len(strSecondForeName), Null)
    adoCommand.Parameters.Append adoParameter

    Set adoParameter = adoCommand.CreateParameter(, adBSTR, adParamInput, Len(strOtherForeNames), Null)
    adoCommand.Parameters.Append adoParameter
    
    Set adoParameter = adoCommand.CreateParameter(, adBSTR, adParamInput, Len(strDoB), Null)
    adoCommand.Parameters.Append adoParameter

    Set adoParameter = adoCommand.CreateParameter(, adBSTR, adParamInput, Len(strPostcode), Null)
    adoCommand.Parameters.Append adoParameter
    
    ' PSC 09/01/01 SYS1750 - Start
    Set adoParameter = adoCommand.CreateParameter(, adBSTR, adParamInput, Len(strHouseName), Null)
    adoCommand.Parameters.Append adoParameter
    
    Set adoParameter = adoCommand.CreateParameter(, adBSTR, adParamInput, Len(strHouseNumber), Null)
    adoCommand.Parameters.Append adoParameter

    Set adoParameter = adoCommand.CreateParameter(, adBSTR, adParamInput, Len(strFlatNumber), Null)
    adoCommand.Parameters.Append adoParameter

    Set adoParameter = adoCommand.CreateParameter(, adBSTR, adParamInput, Len(strPhoneNo), Null)
    adoCommand.Parameters.Append adoParameter

    Set adoParameter = adoCommand.CreateParameter(, adBSTR, adParamInput, Len(strChannelId), Null)
    adoCommand.Parameters.Append adoParameter

    Set adoParameter = adoCommand.CreateParameter(, adBSTR, adParamInput, Len(strOtherSysCustNo), Null)
    adoCommand.Parameters.Append adoParameter
    ' PSC 09/01/01 SYS1750 - End
    
    ' AS 04/07/01 CC012 SQL Server stored procedure returns total number of matching records.
    If eDbEngineType = SQLServer Then
        Set adoParameter = adoCommand.CreateParameter(, adInteger, adParamOutput)
        adoCommand.Parameters.Append adoParameter
    End If
    
    adoCommand(0) = intStartRow
    adoCommand(1) = intPageSize
    If Not strSurname = "" Then
        adoCommand(2) = strSurname
    End If
    If Not strForename = "" Then
        adoCommand(3) = strForename
    End If
    If Not strSecondForeName = "" Then
        adoCommand(4) = strSecondForeName
    End If
    If Not strOtherForeNames = "" Then
        adoCommand(5) = strOtherForeNames
    End If
    If Not strDoB = "" Then
        adoCommand(6) = strDoB
    End If
    If Not strPostcode = "" Then
        adoCommand(7) = strPostcode
    End If
    
    ' PSC 09/01/01 SYS1750 - Start
    If Not strHouseName = "" Then
        adoCommand(8) = strHouseName
    End If
    If Not strHouseNumber = "" Then
        adoCommand(9) = strHouseNumber
    End If
    If Not strFlatNumber = "" Then
        adoCommand(10) = strFlatNumber
    End If
    If Not strPhoneNo = "" Then
        adoCommand(11) = strPhoneNo
    End If
    If Not strChannelId = "" Then
        adoCommand(12) = strChannelId
    End If
    If Not strOtherSysCustNo = "" Then
        adoCommand(13) = strOtherSysCustNo
    End If
    ' PSC 09/01/01 SYS1750 - End
#Else
    Set adoParameter = adoCommand.CreateParameter(, adInteger, adParamInput)
    adoCommand.Parameters.Append adoParameter
    adoCommand(0) = intStartRow

    Set adoParameter = adoCommand.CreateParameter(, adInteger, adParamInput)
    adoCommand.Parameters.Append adoParameter
    adoCommand(1) = intPageSize

    Set adoParameter = adoCommand.CreateParameter(, adBSTR, adParamInput, Len(strSurname), strSurname)
    adoCommand.Parameters.Append adoParameter

    Set adoParameter = adoCommand.CreateParameter(, adBSTR, adParamInput, Len(strForename), strForename)
    adoCommand.Parameters.Append adoParameter

    Set adoParameter = adoCommand.CreateParameter(, adBSTR, adParamInput, Len(strSecondForeName), strSecondForeName)
    adoCommand.Parameters.Append adoParameter

    Set adoParameter = adoCommand.CreateParameter(, adBSTR, adParamInput, Len(strOtherForeNames), strOtherForeNames)
    adoCommand.Parameters.Append adoParameter
    
    Set adoParameter = adoCommand.CreateParameter(, adBSTR, adParamInput, Len(strDoB), strDoB)
    adoCommand.Parameters.Append adoParameter

    Set adoParameter = adoCommand.CreateParameter(, adBSTR, adParamInput, Len(strPostcode), strPostcode)
    adoCommand.Parameters.Append adoParameter
    
    ' PSC 09/01/01 SYS1750 - Start
    Set adoParameter = adoCommand.CreateParameter(, adBSTR, adParamInput, Len(strHouseName), strHouseName)
    adoCommand.Parameters.Append adoParameter
    
    Set adoParameter = adoCommand.CreateParameter(, adBSTR, adParamInput, Len(strHouseNumber), strHouseNumber)
    adoCommand.Parameters.Append adoParameter

    Set adoParameter = adoCommand.CreateParameter(, adBSTR, adParamInput, Len(strFlatNumber), strFlatNumber)
    adoCommand.Parameters.Append adoParameter

    Set adoParameter = adoCommand.CreateParameter(, adBSTR, adParamInput, Len(strPhoneNo), strPhoneNo)
    adoCommand.Parameters.Append adoParameter

    Set adoParameter = adoCommand.CreateParameter(, adBSTR, adParamInput, Len(strChannelId), strChannelId)
    adoCommand.Parameters.Append adoParameter

    Set adoParameter = adoCommand.CreateParameter(, adBSTR, adParamInput, Len(strOtherSysCustNo), strOtherSysCustNo)
    adoCommand.Parameters.Append adoParameter
    ' PSC 09/01/01 SYS1750 - End
    
#End If
' AS 30/05/01 CC012 - End
    
    Set rstThisRecSet = adoCommand.Execute
                
' AS 30/05/01 CC012 - Start
#If GENERIC_SQL Then
    ' AS 31/05/01 CC012
    ' SQLOLEDB does not allow disconnected recordsets when calling stored procedures.
    ' Also Recordset.NextRecordset is not supported for disconnected recordsets.
    ' Seems to work OK though for Oracle OLEDB!
    If eDbEngineType = SQLServer Then
        ' AS 04/07/01 CC012 SQL Server stored procedure returns total number of matching records.
        intRows = adoCommand(14)
        If intRows = 0 Then
            objErrAssist.ThrowError TypeName(Me), strFunctionName, oeRecordNotFound
        End If
    ElseIf eDbEngineType = Oracle Then
        Set adoCommand.ActiveConnection = Nothing
        ' DJP SYS1779, disconnect the recordset before closing.
        Set rstThisRecSet.ActiveConnection = Nothing
        adoConnection.Close
        Set adoConnection = Nothing
    End If
#Else
    Set adoCommand.ActiveConnection = Nothing
    ' DJP SYS1779, disconnect the recordset before closing.
    Set rstThisRecSet.ActiveConnection = Nothing
    adoConnection.Close
    Set adoConnection = Nothing
#End If
' AS 30/05/01 CC012 - End

#If GENERIC_SQL Then
    Dim bValidRecordset As Boolean
    bValidRecordset = adoGetValidRecordset(rstThisRecSet)
    
    If eDbEngineType = SQLServer Then
        ' AS 31/05/01 CC012 Now safe to disconnect the recordset.
        Set adoCommand.ActiveConnection = Nothing
        ' DJP SYS1779, disconnect the recordset before closing.
        Set rstThisRecSet.ActiveConnection = Nothing
        adoConnection.Close
        Set adoConnection = Nothing
    End If
    
    If bValidRecordset Then
#End If
        If rstThisRecSet.EOF Then
            ' raise application error to be interpreted by calling object
            objErrAssist.ThrowError TypeName(Me), strFunctionName, oeRecordNotFound
        Else
            GetXmlFromFindList rstThisRecSet, vobjXmlResponse
            
            Set xmlElem = vobjXmlResponse.selectSingleNode("CUSTOMERLIST")
            xmlElem.setAttribute "TOTAL", intRows
            Set xmlElem = Nothing
            
        End If
#If GENERIC_SQL Then
    Else
        objErrAssist.ThrowError TypeName(Me), strFunctionName, oeRecordNotFound
    End If
#End If
    
    rstThisRecSet.Close
    
    Set rstThisRecSet = Nothing
    Set adoCommand = Nothing
        
    Set objADOAssist = Nothing
    Set objSQLAssist = Nothing
    Set objErrAssist = Nothing
    
    Exit Sub

FindCustomerVbErr:
    
    ' Close the recordset if still open
        If Not rstThisRecSet Is Nothing Then
                If rstThisRecSet.State = adStateOpen Then
                        rstThisRecSet.Close
                End If
        End If

    ' Close the connection if still open
    If Not adoConnection Is Nothing Then
        If adoConnection.State = adStateOpen Then
            adoConnection.Close
        End If
    End If

    Set xmlNode = Nothing
    
    Set rstThisRecSet = Nothing
    Set adoCommand = Nothing
    Set adoConnection = Nothing
        
    Set objADOAssist = Nothing
    Set objSQLAssist = Nothing
    Set objErrAssist = Nothing

    '   re-raise error for business object to interpret as appropriate
    Debug.Print Err.Number & ": " & Err.Description
    Err.Raise Err.Number, Err.Source, Err.Description

End Sub

Private Sub GetXmlFromFindList(ByVal vrst As ADODB.Recordset, _
                               ByVal vobjXmlOut As IXMLDOMNode)

    Dim objContext As ObjectContext
    Set objContext = GetObjectContext()

    Dim objXmlCustomerSchemaDoc As New FreeThreadedDOMDocument40
    objXmlCustomerSchemaDoc.validateOnParse = False
    objXmlCustomerSchemaDoc.setProperty "NewParser", True
    Dim objXmlCustomerVersionSchemaDoc As New FreeThreadedDOMDocument40
    objXmlCustomerVersionSchemaDoc.validateOnParse = False
    objXmlCustomerVersionSchemaDoc.setProperty "NewParser", True
    Dim objXmlAddressSchemaDoc As New FreeThreadedDOMDocument40
    objXmlAddressSchemaDoc.validateOnParse = False
    objXmlAddressSchemaDoc.setProperty "NewParser", True
    
    Dim objXmlElem As IXMLDOMElement
    Dim objXmlCustomerListNode As IXMLDOMNode
    Dim objXmlCustomerNode As IXMLDOMNode
    Dim objXmlCustomerVersionNode As IXMLDOMNode
    Dim objXmlCustomerAddressListNode As IXMLDOMNode
    Dim objXmlCustomerAddressNode As IXMLDOMNode
    Dim objXmlAddressNode As IXMLDOMNode
    Dim xmlSourceNode As IXMLDOMElement
    
    Dim objXmlAssist As New XMLAssist
    
    Dim objClassDefs As IomCustClassDef
    Set objClassDefs = New omCustClassDef
    
    Dim objAddressClassDef As IAddressClassDef
    Set objAddressClassDef = New AddressClassDef
    
    Dim objDOAssist As IDOAssist
    If objContext Is Nothing Then
        Set objDOAssist = New DOAssist
    Else
        Set objDOAssist = objContext.CreateInstance(gstrBASE_COMPONENT & ".DOAssist")
    End If
    
    Dim strCustomerFields(4) As Variant
    Dim strCustVersionFields(10) As Variant         ' PSC 12/10/2005 MAR57
    Dim strAddressFields(9) As Variant
    
    Dim intRecord As Integer
    
    strCustomerFields(0) = "CUSTOMERNUMBER"
    strCustomerFields(1) = "OTHERSYSTEMCUSTOMERNUMBER"
    strCustomerFields(2) = "OTHERSYSTEMTYPE"
    strCustomerFields(3) = "CHANNELID"
    
    strCustVersionFields(0) = "CUSTOMERNUMBER"
    strCustVersionFields(1) = "CUSTOMERVERSIONNUMBER"
    strCustVersionFields(2) = "SURNAME"
    strCustVersionFields(3) = "FIRSTFORENAME"
    strCustVersionFields(4) = "SECONDFORENAME"
    strCustVersionFields(5) = "OTHERFORENAMES"
    strCustVersionFields(6) = "DATEOFBIRTH"
    strCustVersionFields(7) = "CORRESPONDENCESALUTATION"
    strCustVersionFields(8) = "MOTHERSMAIDENNAME"
    strCustVersionFields(9) = "CUSTOMERSTATUS"      ' PSC 12/10/2005 MAR57
                
    strAddressFields(0) = "POSTCODE"
    strAddressFields(1) = "FLATNUMBER"
    strAddressFields(2) = "BUILDINGORHOUSENAME"
    strAddressFields(3) = "BUILDINGORHOUSENUMBER"
    strAddressFields(4) = "STREET"
    strAddressFields(5) = "DISTRICT"
    strAddressFields(6) = "TOWN"
    strAddressFields(7) = "COUNTY"
    strAddressFields(8) = "COUNTRY"
    
    Set objXmlCustomerSchemaDoc = objClassDefs.LoadCustomerData
    Set objXmlCustomerVersionSchemaDoc = objClassDefs.LoadCustomerVersionData
    Set objXmlAddressSchemaDoc = objAddressClassDef.LoadAddressData
    
    objXmlAssist.SelectSchemaFields strCustomerFields, objXmlCustomerSchemaDoc.firstChild
    objXmlAssist.SelectSchemaFields strCustVersionFields, objXmlCustomerVersionSchemaDoc.firstChild
    objXmlAssist.SelectSchemaFields strAddressFields, objXmlAddressSchemaDoc.firstChild
    
    ' create CUSTOMERLIST node
    Set objXmlElem = vobjXmlOut.ownerDocument.createElement("CUSTOMERLIST")
    Set objXmlCustomerListNode = vobjXmlOut.appendChild(objXmlElem)
    
    ' SR 08/06/01 : SYS2362 - Append Source element with text as 'Omiga' to customer node
    Set xmlSourceNode = vobjXmlOut.ownerDocument.createElement("SOURCE")
    xmlSourceNode.Text = "Omiga"
    ' SR 08/06/01 - End
    
    Do While Not vrst.EOF

        Set objXmlCustomerNode = _
            objDOAssist.GetXMLFromRecordSet _
                (vrst, objXmlCustomerSchemaDoc, objXmlCustomerListNode)
                
        objXmlCustomerNode.appendChild xmlSourceNode.cloneNode(True) ' SR 08/06/01 : SYS2362

        Set objXmlCustomerVersionNode = _
            objDOAssist.GetXMLFromRecordSet _
                (vrst, objXmlCustomerVersionSchemaDoc, objXmlCustomerNode)

        Set objXmlElem = vobjXmlOut.ownerDocument.createElement("CUSTOMERADDRESS")
        Set objXmlCustomerAddressNode = objXmlCustomerVersionNode.appendChild(objXmlElem)

        Set objXmlAddressNode = _
            objDOAssist.GetXMLFromRecordSet _
                (vrst, objXmlAddressSchemaDoc, objXmlCustomerAddressNode)

        vrst.MoveNext
    
    Loop

    ' clean up
    Set objXmlCustomerSchemaDoc = Nothing
    Set objXmlCustomerVersionSchemaDoc = Nothing
    Set objXmlAddressSchemaDoc = Nothing
    Set objXmlElem = Nothing
    Set objXmlCustomerListNode = Nothing
    Set objXmlCustomerNode = Nothing
    Set objXmlCustomerVersionNode = Nothing
    Set objXmlCustomerAddressListNode = Nothing
    Set objXmlCustomerAddressNode = Nothing
    Set objXmlAddressNode = Nothing
    Set xmlSourceNode = Nothing
    
    Set objClassDefs = Nothing
    Set objAddressClassDef = Nothing
    Set objXmlAssist = Nothing
    Set objDOAssist = Nothing
    
    Set objContext = Nothing

End Sub

Private Sub ICustomerDO_GetCustomerDetails(ByVal vobjRequestNode As IXMLDOMNode, ByVal vobjResponseNode As IXMLDOMNode)
 On Error GoTo GetCustomerDetailsVbErr
    
    Const strFunctionName As String = "ICustomerDO_GetCustomerDetails"

    Dim objContext As ObjectContext
    Set objContext = GetObjectContext()
    
    Dim objClassDefs As IomCustClassDef
    Dim objAddressClassDef As IAddressClassDef
    Set objAddressClassDef = New AddressClassDef
    
    Dim objErrAssist As New ErrAssist
    Dim objSQLAssist As New SQLAssist
    Dim objXmlAssist As New XMLAssist
    
    Dim objXmlSchemaDoc1 As FreeThreadedDOMDocument40
    Dim objXmlSchemaDoc2 As FreeThreadedDOMDocument40
    Dim objXmlParentNode As IXMLDOMNode
    Dim objXmlParentListNode As IXMLDOMNode
    Dim objXMLElement As IXMLDOMElement
    Dim objXmlCustomerVersionNode As IXMLDOMNode

    Dim objDOAssist As DOAssist
    Dim objADOAssist As ADOAssist
    
    Dim conn As New ADODB.Connection
    Dim rst As New ADODB.Recordset
    
    Dim strCustomerNumber As String, strCustomerVersionNumber As String
    Dim strSQL As String
    
    Set objClassDefs = New omCustClassDef

    If objContext Is Nothing Then
        Set objDOAssist = New DOAssist
        Set objADOAssist = New ADOAssist
    Else
        Set objDOAssist = _
            objContext.CreateInstance(gstrBASE_COMPONENT & ".DOAssist")
        Set objADOAssist = _
            objContext.CreateInstance(gstrBASE_COMPONENT & ".ADOAssist")
    End If
    
    strCustomerNumber = _
        objXmlAssist.GetMandatoryElementText(vobjRequestNode, "SEARCH/CUSTOMER/CUSTOMERNUMBER")
    
    strCustomerVersionNumber = _
        objXmlAssist.GetMandatoryElementText(vobjRequestNode, "SEARCH/CUSTOMER/CUSTOMERVERSIONNUMBER")
        
    strCustomerNumber = objSQLAssist.FormatString(strCustomerNumber)
    
    conn.ConnectionString = objADOAssist.GetConnStr
    conn.Open
    
    rst.CursorLocation = adUseClient
    
    ' customer, customerversion
    strSQL = "select * from customer, customerversion " & _
             "where customerversion.customernumber = " & strCustomerNumber & _
             " and customerversion.customerversionnumber = " & strCustomerVersionNumber & _
             " and customer.customernumber = customerversion.customernumber"

    rst.Open strSQL, conn, adOpenStatic, adLockReadOnly, adCmdText
    
    ' disconnect RecordSet
    Set rst.ActiveConnection = Nothing
    
    If rst.EOF Then
        objErrAssist.ThrowError TypeName(Me), strFunctionName, oeRecordNotFound
    End If
    
    rst.MoveFirst
    
    Set objXmlSchemaDoc1 = objClassDefs.LoadCustomerData
            
    Set objXmlParentNode = _
        objDOAssist.GetXMLFromRecordsetEx(rst, objXmlSchemaDoc1, vobjResponseNode)
        
    Set objXmlSchemaDoc1 = Nothing
    Set objXmlSchemaDoc1 = objClassDefs.LoadCustomerVersionData
            
    Set objXmlCustomerVersionNode = _
        objDOAssist.GetXMLFromRecordsetEx(rst, objXmlSchemaDoc1, objXmlParentNode)
        
    AddCustVersionDerivedData objXmlCustomerVersionNode
    
    rst.Close
    Set objXmlSchemaDoc1 = Nothing
    
    ' customeraddress, address
    
    Set objXMLElement = vobjResponseNode.ownerDocument.createElement("CUSTOMERADDRESSLIST")
    Set objXmlParentListNode = objXmlCustomerVersionNode.appendChild(objXMLElement)
    
    strSQL = "select * from customeraddress, address " & _
             "where customeraddress.customernumber = " & strCustomerNumber & _
             " and customeraddress.customerversionnumber = " & strCustomerVersionNumber & _
             " and customeraddress.addresstype in (1, 2)" & _
             " and address.addressguid = customeraddress.addressguid" & _
             " order by customeraddress.addresstype"

    rst.Open strSQL, conn, adOpenStatic, adLockReadOnly, adCmdText
    
    ' disconnect RecordSet
    
    Set rst.ActiveConnection = Nothing
    
    If Not rst.EOF Then
    
        Set objXmlSchemaDoc1 = objClassDefs.LoadCustomerAddressData
        Set objXmlSchemaDoc2 = objAddressClassDef.LoadAddressData
    
        rst.MoveFirst
        
        Do While Not rst.EOF
            
            Set objXmlParentNode = _
                objDOAssist.GetXMLFromRecordsetEx(rst, objXmlSchemaDoc1, objXmlParentListNode)
            
            Set objXmlParentNode = _
                objDOAssist.GetXMLFromRecordsetEx(rst, objXmlSchemaDoc2, objXmlParentNode)
                
            rst.MoveNext
        
        Loop
    
    End If
    
    rst.Close
    Set objXmlSchemaDoc1 = Nothing
    Set objXmlSchemaDoc2 = Nothing
    
    ' customertelephonenumber
    
    Set objXMLElement = vobjResponseNode.ownerDocument.createElement("CUSTOMERTELEPHONENUMBERLIST")
    Set objXmlParentListNode = objXmlCustomerVersionNode.appendChild(objXMLElement)
    
    strSQL = "select * from customertelephonenumber " & _
             "where customernumber = " & strCustomerNumber & _
             " and customerversionnumber = " & strCustomerVersionNumber

    rst.Open strSQL, conn, adOpenStatic, adLockReadOnly, adCmdText
    
    ' disconnect RecordSet
    
    Set rst.ActiveConnection = Nothing
    
    If Not rst.EOF Then
    
        Set objXmlSchemaDoc1 = objClassDefs.LoadCustomerTelephoneData
    
        rst.MoveFirst
        
        Do While Not rst.EOF
            
            Set objXmlParentNode = _
                objDOAssist.GetXMLFromRecordsetEx(rst, objXmlSchemaDoc1, objXmlParentListNode)
                
            rst.MoveNext
        
        Loop
    
    End If
    
    rst.Close
    Set objXmlSchemaDoc1 = Nothing
    
    'BMIDS864  Copy Customer Verification
    
    Set objXMLElement = vobjResponseNode.ownerDocument.createElement("VERIFICATIONLIST")
    Set objXmlParentListNode = objXmlCustomerVersionNode.appendChild(objXMLElement)
    
    strSQL = "select * from verification " & _
             "where customernumber = " & strCustomerNumber & _
             " and customerversionnumber = " & strCustomerVersionNumber

    rst.Open strSQL, conn, adOpenStatic, adLockReadOnly, adCmdText
    
    ' disconnect RecordSet
    
    Set rst.ActiveConnection = Nothing
    
    If Not rst.EOF Then
    
        Set objXmlSchemaDoc1 = objClassDefs.LoadVerificationData
    
        rst.MoveFirst
        
        Do While Not rst.EOF
            
            Set objXmlParentNode = _
                objDOAssist.GetXMLFromRecordsetEx(rst, objXmlSchemaDoc1, objXmlParentListNode)
                
            rst.MoveNext
        
        Loop
    
    End If
    
    rst.Close
    Set objXmlSchemaDoc1 = Nothing
    
    ' areasofinterest
    
    Set objXMLElement = vobjResponseNode.ownerDocument.createElement("AREASOFINTERESTLIST")
    Set objXmlParentListNode = objXmlCustomerVersionNode.appendChild(objXMLElement)
    
    strSQL = "select * from areasofinterest " & _
             "where customernumber = " & strCustomerNumber & _
             " and customerversionnumber = " & strCustomerVersionNumber

    rst.Open strSQL, conn, adOpenStatic, adLockReadOnly, adCmdText
    
    ' disconnect RecordSet
    
    Set rst.ActiveConnection = Nothing
    
    If Not rst.EOF Then
    
        Set objXmlSchemaDoc1 = objClassDefs.LoadAreasOfInterestData
    
        rst.MoveFirst
        
        Do While Not rst.EOF
            
            Set objXmlParentNode = _
                objDOAssist.GetXMLFromRecordsetEx(rst, objXmlSchemaDoc1, objXmlParentListNode)
                
            rst.MoveNext
        
        Loop
    
    End If
    
    rst.Close
    Set objXmlSchemaDoc1 = Nothing
    
    ' HM MAR29 GetCustomerDetails has been extended to include wrap up details
    ' customerwrapupdetails
    
    Set objXMLElement = vobjResponseNode.ownerDocument.createElement("WRAPUP")
    Set objXmlParentListNode = objXmlCustomerVersionNode.appendChild(objXMLElement)
    
    strSQL = "select * from CustomerWrapUpDetails " & _
             "where customernumber = " & strCustomerNumber & _
             " and customerversionnumber = " & strCustomerVersionNumber

    rst.Open strSQL, conn, adOpenStatic, adLockReadOnly, adCmdText
    
    ' disconnect RecordSet
    
    Set rst.ActiveConnection = Nothing
    
    If Not rst.EOF Then
    
        Set objXmlSchemaDoc1 = objClassDefs.LoadWrapUpDetailsData
    
        rst.MoveFirst
        
        Do While Not rst.EOF
            
            Set objXmlParentNode = _
                objDOAssist.GetXMLFromRecordsetEx(rst, objXmlSchemaDoc1, objXmlParentListNode)
                
            rst.MoveNext
        
        Loop
    
    End If
    
    rst.Close
    Set objXmlSchemaDoc1 = Nothing

    Debug.Print vobjResponseNode.xml

    Set rst = Nothing
        conn.Close
    Set conn = Nothing
    
    Set objClassDefs = Nothing
    Set objAddressClassDef = Nothing
    Set objErrAssist = Nothing
    Set objSQLAssist = Nothing
    Set objXmlAssist = Nothing
    Set objDOAssist = Nothing

    Set objADOAssist = Nothing
    Set objXmlParentNode = Nothing
    Set objXmlParentListNode = Nothing
    Set objXMLElement = Nothing
    Set objXmlCustomerVersionNode = Nothing
    Set objXmlSchemaDoc1 = Nothing
    Set objXmlSchemaDoc2 = Nothing
    
    Exit Sub
    
GetCustomerDetailsVbErr:
    
    ' Close the recordset if still open
        If Not rst Is Nothing Then
                If rst.State = adStateOpen Then
                        rst.Close
                End If
        End If

    ' Close the connection if still open
        If Not conn Is Nothing Then
                If conn.State = adStateOpen Then
                        conn.Close
                End If
        End If

    Set rst = Nothing
        Set conn = Nothing
        Set objClassDefs = Nothing
    Set objAddressClassDef = Nothing
    Set objErrAssist = Nothing
    Set objSQLAssist = Nothing
    Set objXmlAssist = Nothing
    Set objDOAssist = Nothing

    Set objADOAssist = Nothing
    Set objXmlParentNode = Nothing
    Set objXmlParentListNode = Nothing
    Set objXMLElement = Nothing
    Set objXmlCustomerVersionNode = Nothing
    Set objXmlSchemaDoc1 = Nothing
    Set objXmlSchemaDoc2 = Nothing

    ' re-raise the error back up the calling chain
    Debug.Print Err.Number & ": " & Err.Description
    Err.Raise Err.Number, Err.Source, Err.Description, Err.HelpFile, Err.HelpContext

End Sub

Private Sub ICustomerDO_CreateCustomerLock( _
    ByVal vobjXmlRequestNode As IXMLDOMNode, _
    ByVal vobjXmlResponseNode As IXMLDOMNode, _
    Optional ByVal vstrLockType As String = "On")
    
    Dim objContext As ObjectContext
    Set objContext = GetObjectContext()
    
    Dim objADOAssist As ADOAssist
    Dim objSQLAssist As SQLAssist
    Set objSQLAssist = New SQLAssist

    If objContext Is Nothing Then
        Set objADOAssist = New ADOAssist
    Else
        Set objADOAssist = _
            objContext.CreateInstance(gstrBASE_COMPONENT & ".ADOAssist")
    End If
    
    Dim objXmlAssist As New XMLAssist
    
    Dim strSQL As String, strFields As String, strValues As String
    Dim strCustomerNumber As String, strUserId As String
    Dim strUnitID As String, strMachineId As String
    
    ' get customerlock parameters
    strCustomerNumber = objXmlAssist.GetMandatoryElementText(vobjXmlRequestNode, ".//CUSTOMERNUMBER")
    strUserId = objXmlAssist.GetAttributeFromNode(vobjXmlRequestNode, "USERID")
    strUnitID = objXmlAssist.GetAttributeFromNode(vobjXmlRequestNode, "UNITID")
    strMachineId = objXmlAssist.GetAttributeFromNode(vobjXmlRequestNode, "MACHINEID")
    
    strFields = "CUSTOMERNUMBER"
    strValues = objSQLAssist.FormatString(strCustomerNumber)
    
    If Len(strUserId) > 0 Then
        strFields = strFields & ", USERID"
        strValues = strValues & ", " & objSQLAssist.FormatString(strUserId)
    End If
    
    If Len(strUnitID) > 0 Then
        strFields = strFields & ", UNITID"
        strValues = strValues & ", " & objSQLAssist.FormatString(strUnitID)
    End If
    
    If Len(strMachineId) > 0 Then
        strFields = strFields & ", MACHINEID"
        strValues = strValues & ", " & objSQLAssist.FormatString(strMachineId)
    End If
    
    strFields = strFields & ", TYPEOFLOCK"
    strValues = strValues & ", " & objSQLAssist.FormatString(vstrLockType)
    
    strFields = strFields & ", LOCKDATE"
    strValues = strValues & ", " & objSQLAssist.FormatDate(Format(Now, "dd/mm/yyyy hh:mm:ss"))
    
    strSQL = "INSERT INTO CUSTOMERLOCK (" & strFields & ")" & _
                " VALUES (" & strValues & ")"

    objADOAssist.executeSQLCommand strSQL
    
    Set objXmlAssist = Nothing
    Set objADOAssist = Nothing
    Set objSQLAssist = Nothing
    Set objContext = Nothing

End Sub

Private Sub ICustomerDO_DeleteCustomerLock(ByVal vobjXmlRequest As IXMLDOMNode, ByVal vobjXmlResponse As IXMLDOMNode)
    
    Dim objContext As ObjectContext
    Set objContext = GetObjectContext()
    
    Dim objADOAssist As ADOAssist
    Dim objSQLAssist As SQLAssist
    Set objSQLAssist = New SQLAssist

    If objContext Is Nothing Then
        Set objADOAssist = New ADOAssist
    Else
        Set objADOAssist = _
            objContext.CreateInstance(gstrBASE_COMPONENT & ".ADOAssist")
    End If
    
    Dim objXmlAssist As New XMLAssist

    Dim strSQL As String, strWhere As String
    Dim strCustomerNumber As String, strUserId As String
    
    ' get customerlock parameters
    strCustomerNumber = objXmlAssist.GetMandatoryElementText(vobjXmlRequest, "//CUSTOMERNUMBER")
    strUserId = objXmlAssist.GetAttributeFromNode(vobjXmlRequest, "USERID")
    
    strWhere = "CUSTOMERNUMBER=" & objSQLAssist.FormatString(strCustomerNumber)
    
    If Len(strUserId) > 0 Then
        strWhere = strWhere & " AND USERID=" & objSQLAssist.FormatString(strUserId)
    End If
    
    strSQL = "DELETE FROM CUSTOMERLOCK WHERE " & strWhere
    
    objADOAssist.executeSQLCommand strSQL
    
    Set objXmlAssist = Nothing
    Set objADOAssist = Nothing
    Set objSQLAssist = Nothing
    Set objContext = Nothing

End Sub

Private Function ICustomerDO_GetNextCustomerNumber() As Long
On Error GoTo GetNextCustomerNumberVbErr
    
    Dim objContext As ObjectContext
    Set objContext = GetObjectContext()
    
    Const strFunctionName As String = "ICustomerDO_GetNextCustomerNumber"
    
    Dim objADOAssist As ADOAssist

    If objContext Is Nothing Then
        Set objADOAssist = New ADOAssist
    Else
        Set objADOAssist = _
            objContext.CreateInstance(gstrBASE_COMPONENT & ".ADOAssist")
    End If
    
    Dim objErrAssist As New ErrAssist
    
    Dim adoConnection As ADODB.Connection
    Dim adoCommand As ADODB.Command
    
    #If GENERIC_SQL Then 'DM SYS1946 SQLServer port Sequences change
        Dim adoParameterCustomerNumber As ADODB.Parameter
        Const ProcName = "USP_GETNEXTCUSTOMERNUMBER"
        Const OutSize = 12
    #Else
        Dim adoRecSet As ADODB.Recordset
    #End If
    Set adoConnection = New ADODB.Connection
    With adoConnection
        .ConnectionString = objADOAssist.GetConnStr
        #If GENERIC_SQL Then 'DM SYS1946 SQLServer port Sequences change
            'Nothing to do in here
        #Else
            .CursorLocation = adUseClient
        #End If
        .Open
    End With
    
    Set adoCommand = New ADODB.Command
    With adoCommand
        Set .ActiveConnection = adoConnection
        #If GENERIC_SQL Then 'DM SYS1946 SQLServer port Sequences change
            .CommandType = adCmdStoredProc
            .CommandText = ProcName
            'DM SYS2511 Changed parameter type to adVarChar
            Set adoParameterCustomerNumber = adoCommand.CreateParameter("p_NextNumber", adVarChar, adParamOutput, 12)
            adoCommand.Parameters.Append adoParameterCustomerNumber
        #Else
            .CommandType = adCmdText
            .CommandText = "SELECT SEQ_CUSTOMERNO.NextVal FROM DUAL"
        #End If
        .Prepared = True
    End With
          
    'IK_17/05/2006_EP561
    If IsCheckDigitRequired Then
        ' Keep getting numbers and calculating the check digit until the
        ' number returned produces a valid Check Digit
        #If GENERIC_SQL Then 'DM SYS1946 SQLServer port Sequences change
            Do
                adoCommand.Execute
            Loop Until Len(GenerateCheckDigit(Format$(adoParameterCustomerNumber.Value, "0000000"))) > 0
        #Else
        
            Do
                Set adoRecSet = adoCommand.Execute
                If adoRecSet.EOF Then
                    objErrAssist.ThrowError TypeName(Me), strFunctionName, oeRecordNotFound
                End If
            Loop Until Len(GenerateCheckDigit(Format$(adoRecSet.Fields(0).Value, "0000000"))) > 0
    
        #End If
        
    Else
        
        #If GENERIC_SQL Then 'DM SYS1946 SQLServer port Sequences change
            adoCommand.Execute
        #Else
            Set adoRecSet = adoCommand.Execute
            If adoRecSet.EOF Then
                objErrAssist.ThrowError TypeName(Me), strFunctionName, oeRecordNotFound
            End If
        #End If
    
    End If
    'IK_17/05/2006_EP561_ends

    #If GENERIC_SQL Then 'DM SYS1946 SQLServer port Sequences change
        ' Nothing required in here
    #Else
        Set adoRecSet.ActiveConnection = Nothing    ' CL SYS1695
    #End If
    Set adoCommand.ActiveConnection = Nothing
    adoConnection.Close
    Set adoConnection = Nothing
       
    #If GENERIC_SQL Then 'DM SYS1946 SQLServer port Sequences change
        ICustomerDO_GetNextCustomerNumber = adoParameterCustomerNumber.Value
    #Else
        ICustomerDO_GetNextCustomerNumber = adoRecSet.Fields(0).Value
        adoRecSet.Close
    #End If
    
    Set adoConnection = Nothing
    Set adoCommand = Nothing
    #If GENERIC_SQL Then
        'Nothing to do in here
    #Else
        Set adoRecSet = Nothing
    #End If
    Set objADOAssist = Nothing
    Set objErrAssist = Nothing
    
    Exit Function
    
GetNextCustomerNumberVbErr:
    
    ' Close the recordset if still open
    #If GENERIC_SQL Then 'DM SYS1946 SQLServer port Sequences change
        'Nothing to do in here
    #Else
        If Not adoRecSet Is Nothing Then
                If adoRecSet.State = adStateOpen Then
                        adoRecSet.Close
                End If
        End If
    #End If

    ' Close the connection if still open
        If Not adoConnection Is Nothing Then
                If adoConnection.State = adStateOpen Then
                        adoConnection.Close
                End If
        End If

    Set adoConnection = Nothing
    Set adoCommand = Nothing
    #If GENERIC_SQL Then 'DM SYS1946 SQLServer port Sequences change
        ' Nothing to do in here
    #Else
        Set adoRecSet = Nothing
    #End If
    Set objADOAssist = Nothing
    Set objErrAssist = Nothing

    ' re-raise the error back up the calling chain
    Debug.Print Err.Number & ": " & Err.Description
    Err.Raise Err.Number, Err.Source, Err.Description, Err.HelpFile, Err.HelpContext

End Function

Private Sub ICustomerDO_CreateCustomerDetails( _
    ByVal vobjRequestNode As IXMLDOMNode, _
    ByVal vobjResponseNode As IXMLDOMNode, _
    ByVal vstrCustomerNumber As String, _
    ByVal vstrCustomerVersionNumber As String, _
    Optional ByVal vblnCreateNewCustomer As Boolean = True)

'BMIDS00004 MDC 13/05/2002
'New optional parameter to indicate if a Customer record needs to be created.
'If creating a new customer version for an existing customer, the existing
'customer record should be used.

On Error GoTo CreateCustomerDetailsVbErr
    
    Dim objContext As ObjectContext
    Set objContext = GetObjectContext()
    
    Dim objXmlCustomerNode As IXMLDOMNode
    Dim objXmlCustomerVersionNode As IXMLDOMNode
    Dim objXmlCustomerAddressNode As IXMLDOMNode
    Dim objXmlAddressNode As IXMLDOMNode
    Dim objXmlNodeList As IXMLDOMNodeList
    Dim objXmlNode As IXMLDOMNode
    Dim xmlChannelIdElem As IXMLDOMElement
    
    Dim objXmlSchemaDoc As FreeThreadedDOMDocument40
    Dim objXmlCustomerAddressSchema As FreeThreadedDOMDocument40
    Dim objXmlAddressSchema As FreeThreadedDOMDocument40
    
    Dim objXmlAssist As New XMLAssist
    
    Dim strAddressGuid As String
    Dim IntCount As Integer
    
    'BMIDS00004 MDC 13/05/2002
    Dim xmlCustomerVersionNode As IXMLDOMNode
    Dim xmlOut As FreeThreadedDOMDocument40
    'BMIDS00004 MDC 13/05/2002 - End
    
    Dim objClassDefs As IomCustClassDef
    Set objClassDefs = New omCustClassDef
    
    Dim objAddressClassDef As IAddressClassDef
    Set objAddressClassDef = New AddressClassDef
    
    Dim objDOAssist As DOAssist
    If objContext Is Nothing Then
        Set objDOAssist = New DOAssist
    Else
        Set objDOAssist = _
            objContext.CreateInstance(gstrBASE_COMPONENT & ".DOAssist")
    End If
        
    ' proliferate keys
    Set objXmlCustomerNode = objXmlAssist.GetMandatoryNode(vobjRequestNode, "CREATE/CUSTOMER")
    objXmlAssist.SetMandatoryChildText objXmlCustomerNode, "CUSTOMERNUMBER", vstrCustomerNumber

    
    Set objXmlCustomerVersionNode = objXmlAssist.GetMandatoryNode(objXmlCustomerNode, "CUSTOMERVERSION")
    'BMIDS00004 MDC 13/05/2002
    If Len(Trim$(vstrCustomerVersionNumber)) = 0 Then
        Set objXmlSchemaDoc = objClassDefs.LoadCustomerVersionData
        objDOAssist.GenerateSequenceNumberEx objXmlCustomerVersionNode, objXmlSchemaDoc, "CUSTOMERVERSIONNUMBER"
        vstrCustomerVersionNumber = objXmlAssist.GetElementText(objXmlCustomerVersionNode, "CUSTOMERVERSIONNUMBER")
    End If
    'BMIDS00004 MDC 13/05/2002 - End
    
    objXmlAssist.SetMandatoryChildText objXmlCustomerVersionNode, "CUSTOMERNUMBER", vstrCustomerNumber
    objXmlAssist.SetMandatoryChildText objXmlCustomerVersionNode, "CUSTOMERVERSIONNUMBER", vstrCustomerVersionNumber
    
    '   Create tag for Channel Id and assign the value mentioned in Request tag
    ' SR - AQR: SYS/0023
    Dim strChannelId As String
    Set xmlChannelIdElem = objXmlCustomerNode.selectSingleNode(".//CHANNELID")
    If xmlChannelIdElem Is Nothing Then
        Set xmlChannelIdElem = vobjRequestNode.ownerDocument.createElement("CHANNELID")
        objXmlCustomerNode.appendChild xmlChannelIdElem
    End If
    xmlChannelIdElem.Text = objXmlAssist.GetAttributeFromNode(vobjRequestNode, "CHANNELID")
    
    ' add salutation to CustomerVersion
    AddSalutationToRequest objXmlCustomerVersionNode
    
    Set objXmlNodeList = objXmlCustomerVersionNode.selectNodes("CUSTOMERADDRESSLIST/CUSTOMERADDRESS")
    For IntCount = 0 To (objXmlNodeList.length - 1)
        
        Set objXmlCustomerAddressNode = objXmlNodeList.Item(IntCount)
        Set objXmlAddressNode = objXmlAssist.GetMandatoryNode(objXmlCustomerAddressNode, "ADDRESS")
        
        strAddressGuid = CreateGUID()
    
        objXmlAssist.SetMandatoryChildText objXmlCustomerAddressNode, "CUSTOMERNUMBER", vstrCustomerNumber
        objXmlAssist.SetMandatoryChildText objXmlCustomerAddressNode, "CUSTOMERVERSIONNUMBER", vstrCustomerVersionNumber
    
    Next
    
    Set objXmlNodeList = objXmlCustomerVersionNode.selectNodes("CUSTOMERTELEPHONENUMBERLIST/CUSTOMERTELEPHONENUMBER")
    For IntCount = 0 To (objXmlNodeList.length - 1)
        
        Set objXmlNode = objXmlNodeList.Item(IntCount)
    
        objXmlAssist.SetMandatoryChildText objXmlNode, "CUSTOMERNUMBER", vstrCustomerNumber
        objXmlAssist.SetMandatoryChildText objXmlNode, "CUSTOMERVERSIONNUMBER", vstrCustomerVersionNumber
        objXmlAssist.SetMandatoryChildText objXmlNode, "TELEPHONESEQUENCENUMBER", Format(IntCount + 1)
    
    Next
    
    ' *** BMIDS864  Save Verification
    Set objXmlNodeList = objXmlCustomerVersionNode.selectNodes("VERIFICATIONLIST/VERIFICATION")
    For IntCount = 0 To (objXmlNodeList.length - 1)
        
        Set objXmlNode = objXmlNodeList.Item(IntCount)
    
        objXmlAssist.SetMandatoryChildText objXmlNode, "CUSTOMERNUMBER", vstrCustomerNumber
        objXmlAssist.SetMandatoryChildText objXmlNode, "CUSTOMERVERSIONNUMBER", vstrCustomerVersionNumber
        objXmlAssist.SetMandatoryChildText objXmlNode, "VERIFICATIONSEQUENCENUMBER", Format(IntCount + 1)
    
    Next
    ' *** BMIDS864 End
      
    Set objXmlNodeList = objXmlCustomerVersionNode.selectNodes("AREASOFINTERESTLIST/AREASOFINTEREST")
    For IntCount = 0 To (objXmlNodeList.length - 1)
        
        Set objXmlNode = objXmlNodeList.Item(IntCount)
    
        objXmlAssist.SetMandatoryChildText objXmlNode, "CUSTOMERNUMBER", vstrCustomerNumber
        objXmlAssist.SetMandatoryChildText objXmlNode, "CUSTOMERVERSIONNUMBER", vstrCustomerVersionNumber
    
    Next
    
    ' create records
    
    'BMIDS00004 MDC 13/05/2002 - Create new Customer record if necessary
    If vblnCreateNewCustomer Then
        Set objXmlSchemaDoc = objClassDefs.LoadCustomerData
        objDOAssist.CreateEx objXmlCustomerNode, objXmlSchemaDoc
    End If
    'BMIDS00004 MDC 13/05/2002 - End

    Set objXmlSchemaDoc = objClassDefs.LoadCustomerVersionData
    objDOAssist.CreateEx objXmlCustomerVersionNode, objXmlSchemaDoc
    
    Set objXmlCustomerAddressSchema = objClassDefs.LoadCustomerAddressData
    Set objXmlAddressSchema = objAddressClassDef.LoadAddressData
    
    ' customeraddress / address pairs
    Set objXmlNode = objXmlCustomerVersionNode.selectSingleNode("CUSTOMERADDRESSLIST")
    If Not objXmlNode Is Nothing Then
        WriteCustomerAddresses objDOAssist, objXmlNode
    End If
    
    Set objXmlNode = objXmlCustomerVersionNode.selectSingleNode("CUSTOMERTELEPHONENUMBERLIST")
    If Not objXmlNode Is Nothing Then
        WriteCustomerTelephones vstrCustomerNumber, vstrCustomerVersionNumber, objXmlNode
    End If
    
    ' BMIDS864  Save verification
    Set objXmlNode = objXmlCustomerVersionNode.selectSingleNode("VERIFICATIONLIST")
    If Not objXmlNode Is Nothing Then
        WriteCustomerVerifications vstrCustomerNumber, vstrCustomerVersionNumber, objXmlNode
    End If
    ' BMIDS864  End
    
    Set objXmlNode = objXmlCustomerVersionNode.selectSingleNode("AREASOFINTERESTLIST")
    If Not objXmlNode Is Nothing Then
        WriteAreasOfInterest vstrCustomerNumber, vstrCustomerVersionNumber, objXmlNode
    End If
    
    'BMIDS00004 MDC 13/05/2002
    Set xmlOut = New FreeThreadedDOMDocument40
    xmlOut.validateOnParse = False
    xmlOut.setProperty "NewParser", True
    Set xmlCustomerVersionNode = xmlOut.createElement("CUSTOMERVERSIONNUMBER")
    xmlCustomerVersionNode.Text = vstrCustomerVersionNumber
    vobjResponseNode.appendChild xmlCustomerVersionNode
    'BMIDS00004 MDC 13/05/2002 - End
    
    If Not objContext Is Nothing Then
        objContext.SetComplete
    End If
    
    ' clean up
    
    Set objXmlAssist = Nothing
    Set objDOAssist = Nothing
    Set objClassDefs = Nothing
    Set objAddressClassDef = Nothing
    
    Set objXmlSchemaDoc = Nothing
    Set objXmlCustomerAddressSchema = Nothing
    Set objXmlAddressSchema = Nothing
    Set objXmlCustomerNode = Nothing
    Set objXmlCustomerVersionNode = Nothing
    Set objXmlCustomerAddressNode = Nothing
    Set objXmlAddressNode = Nothing
    Set objXmlNodeList = Nothing
    Set objXmlNode = Nothing
    Set xmlChannelIdElem = Nothing
    'BMIDS00004 MDC 13/05/2002
    Set xmlOut = Nothing
    Set xmlCustomerVersionNode = Nothing
    'BMIDS00004 MDC 13/05/2002 - End
    
    Set objContext = Nothing

    Exit Sub
    
CreateCustomerDetailsVbErr:
    
    If Not objContext Is Nothing Then
        objContext.SetComplete
    End If
    
    Set objXmlAssist = Nothing
    Set objDOAssist = Nothing
    Set objClassDefs = Nothing
    Set objAddressClassDef = Nothing
    
    Set objXmlSchemaDoc = Nothing
    Set objXmlCustomerAddressSchema = Nothing
    Set objXmlAddressSchema = Nothing
    Set objXmlCustomerNode = Nothing
    Set objXmlCustomerVersionNode = Nothing
    Set objXmlCustomerAddressNode = Nothing
    Set objXmlAddressNode = Nothing
    Set objXmlNodeList = Nothing
    Set objXmlNode = Nothing
    
    Set objContext = Nothing

    ' re-raise the error back up the calling chain
    Debug.Print Err.Number & ": " & Err.Description
    Err.Raise Err.Number, Err.Source, Err.Description, Err.HelpFile, Err.HelpContext

End Sub

Private Sub ICustomerDO_UpdateCustomerDetails(ByVal vobjXmlRequest As IXMLDOMNode, ByVal vobjXmlResponse As IXMLDOMNode)
On Error GoTo UpdateCustomerDetailsVbErr

    Const strFunctionName As String = "ICustomerDO_UpdateCustomerDetails"

    Dim objContext As ObjectContext
    Set objContext = GetObjectContext()
    
    Dim objXmlCustomerNode As IXMLDOMNode
    Dim objXmlCustomerVersionNode As IXMLDOMNode
    Dim objXmlCustomerAddressNode As IXMLDOMNode
    Dim objXmlAddressNode As IXMLDOMNode
    Dim objXmlNodeList As IXMLDOMNodeList
    Dim objXmlNode As IXMLDOMNode
    
    Dim objXmlSchemaDoc As FreeThreadedDOMDocument40
    
    Dim objClassDefs As IomCustClassDef
    Set objClassDefs = New omCustClassDef
    
    Dim objXmlAssist As New XMLAssist
    
    Dim strCustomerNumber As String, strCustomerVersionNumber As String
    Dim IntCount As Integer
    
    Dim objDOAssist As DOAssist
    If objContext Is Nothing Then
        Set objDOAssist = New DOAssist
    Else
        Set objDOAssist = _
            objContext.CreateInstance(gstrBASE_COMPONENT & ".DOAssist")
    End If

    ' Get the keys

    Set objXmlCustomerNode = objXmlAssist.GetMandatoryNode(vobjXmlRequest, "UPDATE/CUSTOMER")
    
    strCustomerNumber = _
        objXmlAssist.GetMandatoryElementText(objXmlCustomerNode, "CUSTOMERNUMBER")

    strCustomerVersionNumber = _
        objXmlAssist.GetMandatoryElementText(objXmlCustomerNode, "CUSTOMERVERSION/CUSTOMERVERSIONNUMBER")
    
    ' proliferate keys

    Set objXmlCustomerVersionNode = objXmlAssist.GetMandatoryNode(objXmlCustomerNode, "CUSTOMERVERSION")
    objXmlAssist.SetMandatoryChildText objXmlCustomerVersionNode, "CUSTOMERNUMBER", strCustomerNumber
    objXmlAssist.SetMandatoryChildText objXmlCustomerVersionNode, "CUSTOMERVERSIONNUMBER", strCustomerVersionNumber
    
    ' add salutation to CustomerVersion
    AddSalutationToRequest objXmlCustomerVersionNode
    
    Set objXmlNodeList = objXmlCustomerVersionNode.selectNodes("CUSTOMERADDRESSLIST/CUSTOMERADDRESS")
        
    For IntCount = 0 To (objXmlNodeList.length - 1)
        
        Set objXmlNode = objXmlNodeList.Item(IntCount)
    
        objXmlAssist.SetMandatoryChildText objXmlNode, "CUSTOMERNUMBER", strCustomerNumber
        objXmlAssist.SetMandatoryChildText objXmlNode, "CUSTOMERVERSIONNUMBER", strCustomerVersionNumber
    
    Next
    
    Set objXmlNodeList = objXmlCustomerVersionNode.selectNodes("CUSTOMERTELEPHONENUMBERLIST/CUSTOMERTELEPHONENUMBER")
    For IntCount = 0 To (objXmlNodeList.length - 1)
        
        Set objXmlNode = objXmlNodeList.Item(IntCount)
    
        objXmlAssist.SetMandatoryChildText objXmlNode, "CUSTOMERNUMBER", strCustomerNumber
        objXmlAssist.SetMandatoryChildText objXmlNode, "CUSTOMERVERSIONNUMBER", strCustomerVersionNumber
        objXmlAssist.SetMandatoryChildText objXmlNode, "TELEPHONESEQUENCENUMBER", Format(IntCount + 1)
    
    Next
    
    Set objXmlNodeList = objXmlCustomerVersionNode.selectNodes("AREASOFINTERESTLIST/AREASOFINTEREST")
    For IntCount = 0 To (objXmlNodeList.length - 1)
        
        Set objXmlNode = objXmlNodeList.Item(IntCount)
    
        objXmlAssist.SetMandatoryChildText objXmlNode, "CUSTOMERNUMBER", strCustomerNumber
        objXmlAssist.SetMandatoryChildText objXmlNode, "CUSTOMERVERSIONNUMBER", strCustomerVersionNumber
    
    Next
    
    ' update records
    ' SR - AQR/SYS0220: Check whether any non-primarykey is mentioned in XML, If yes do the updation
    ' else skip this step.
    Dim blnCustomerUpdationReqd As Boolean
    
    Set objXmlSchemaDoc = objClassDefs.LoadCustomerData
    blnCustomerUpdationReqd = Not objXmlAssist.AreAllNonKeyValuesNullEx(objXmlCustomerNode, objXmlSchemaDoc)
    If blnCustomerUpdationReqd Then
        objDOAssist.UpdateEx objXmlCustomerNode, objXmlSchemaDoc
    End If
    Set objXmlSchemaDoc = objClassDefs.LoadCustomerVersionData
    objDOAssist.UpdateEx objXmlCustomerVersionNode, objXmlSchemaDoc
    
    ' customeraddress / address pairs
    Set objXmlNode = objXmlCustomerVersionNode.selectSingleNode("CUSTOMERADDRESSLIST")
    If Not objXmlNode Is Nothing Then
        WriteCustomerAddresses objDOAssist, objXmlNode
    End If
    
    Set objXmlNode = objXmlCustomerVersionNode.selectSingleNode("CUSTOMERTELEPHONENUMBERLIST")
    If Not objXmlNode Is Nothing Then
        WriteCustomerTelephones strCustomerNumber, strCustomerVersionNumber, objXmlNode
    End If
    
    Set objXmlNode = objXmlCustomerVersionNode.selectSingleNode("AREASOFINTERESTLIST")
    If Not objXmlNode Is Nothing Then
        WriteAreasOfInterest strCustomerNumber, strCustomerVersionNumber, objXmlNode
    End If
    
    'clean up
    
    If Not objContext Is Nothing Then
        objContext.SetComplete
    End If
    
    Set objXmlSchemaDoc = Nothing
    Set objXmlCustomerNode = Nothing
    Set objXmlCustomerVersionNode = Nothing
    Set objXmlCustomerAddressNode = Nothing
    Set objXmlAddressNode = Nothing
    Set objXmlNodeList = Nothing
    Set objXmlNode = Nothing
    
    Set objXmlAssist = Nothing
    Set objDOAssist = Nothing
    Set objClassDefs = Nothing
    
    Set objContext = Nothing
    
    Exit Sub
    
UpdateCustomerDetailsVbErr:
    
    'clean up
    
    If Not objContext Is Nothing Then
        objContext.SetComplete
    End If
    
    Set objXmlSchemaDoc = Nothing
    Set objXmlCustomerNode = Nothing
    Set objXmlCustomerVersionNode = Nothing
    Set objXmlCustomerAddressNode = Nothing
    Set objXmlAddressNode = Nothing
    Set objXmlNodeList = Nothing
    Set objXmlNode = Nothing
    
    Set objXmlAssist = Nothing
    Set objDOAssist = Nothing
    Set objClassDefs = Nothing
    
    Set objContext = Nothing

    ' re-raise the error back up the calling chain
    Debug.Print Err.Number & ": " & Err.Description
    Err.Raise Err.Number, Err.Source, Err.Description, Err.HelpFile, Err.HelpContext

End Sub

Private Sub WriteCustomerAddresses(ByVal vobjDOAssist As DOAssist, _
                                   ByVal vobjCustomerAddressList As IXMLDOMNode)

On Error GoTo WriteCustomerAddressesVbErr
    
    Dim objContext As ObjectContext
    Set objContext = GetObjectContext()
    
    Dim objXmlCustomerAddressSchema As FreeThreadedDOMDocument40
    Dim objXmlAddressSchema As FreeThreadedDOMDocument40
    Dim objXmlCustomerAddressNode As IXMLDOMNode
    Dim objXmlAddressNode As IXMLDOMNode
    Dim objXmlNodeList As IXMLDOMNodeList
    Dim objXmlNode As IXMLDOMNode
    
    Dim objXmlAssist As New XMLAssist
    
    Dim objClassDefs As IomCustClassDef
    Set objClassDefs = New omCustClassDef
    
    Dim objAddressClassDef As IAddressClassDef
    Set objAddressClassDef = New AddressClassDef
    
    Dim objAddressDO As IAddressDO
    Set objAddressDO = New AddressDO
    
    'SG 30/05/02 SYS4767 START
    Dim objCustomerDO As ICustomerDO
    Set objCustomerDO = New CustomerDO
    'SG 30/05/02 SYS4767 END
    
    Dim xmlLastAmendedDate As IXMLDOMNode ' EP399
    
    Dim objADOAssist As ADOAssist
    Dim objSQLAssist As SQLAssist
    Set objSQLAssist = New SQLAssist

    If objContext Is Nothing Then
        Set objADOAssist = New ADOAssist
    Else
        Set objADOAssist = _
            objContext.CreateInstance(gstrBASE_COMPONENT & ".ADOAssist")
    End If
    
    'SG 30/05/02 SYS4767 START
    Dim objIADOAssist  As IADOAssist
    Set objIADOAssist = New ADOAssist
    'SG 30/05/02 SYS4767 END
    
    Dim strSQL As String, strWhere As String
    Dim strCustomerNumber As String, strCustomerVersionNumber As String
    Dim strAddressSeq As String, strAddressGuid As String
    
    Dim IntCount As Integer, intNextSeq As Integer
   
    'SG 30/05/02 SYS4767 START
    'Dim blnEmptyAddress As Boolean, blnExistingAddress As Boolean
    Dim blnEmptyAddress As Boolean, blnExistingAddress As Boolean, blnExistingCustomerAddress As Boolean
    Dim strTemp As String, strExistingAddressGUID As String
    'SG 30/05/02 SYS4767 END
        
    Set objXmlNodeList = vobjCustomerAddressList.selectNodes("CUSTOMERADDRESS")
    
    If objXmlNodeList.length > 0 Then
    
        strCustomerNumber = objXmlAssist.GetMandatoryElementText(objXmlNodeList.Item(0), "CUSTOMERNUMBER")
        strCustomerVersionNumber = objXmlAssist.GetMandatoryElementText(objXmlNodeList.Item(0), "CUSTOMERVERSIONNUMBER")
        
        intNextSeq = GetHighestAddressSeqNo(strCustomerNumber, _
                                            strCustomerVersionNumber, _
                                            objADOAssist, objSQLAssist) + 1
        
        Set objXmlCustomerAddressSchema = objClassDefs.LoadCustomerAddressData
        Set objXmlAddressSchema = objAddressClassDef.LoadAddressData

        For IntCount = 0 To (objXmlNodeList.length - 1)
        
            blnEmptyAddress = False
            blnExistingAddress = False
            
            'SG 30/05/02 SYS4767
            blnExistingCustomerAddress = False
                        
            Set objXmlCustomerAddressNode = objXmlNodeList.Item(IntCount)
            Set objXmlAddressNode = objXmlAssist.GetMandatoryNode(objXmlCustomerAddressNode, "ADDRESS")
            
            Set objXmlNode = objXmlAddressNode.selectSingleNode("ADDRESSGUID")
            
            If Not objXmlNode Is Nothing Then
                
                strAddressGuid = objXmlNode.Text
                
                If Len(strAddressGuid) > 0 Then
                    blnExistingAddress = True
                End If
                
            End If
                
            blnEmptyAddress = objAddressDO.IsEmptyAddress(objXmlAddressNode)
            
            If blnExistingAddress Then
            
                strAddressSeq = objXmlAssist.GetMandatoryElementText(objXmlCustomerAddressNode, "CUSTOMERADDRESSSEQUENCENUMBER")
            
                If blnEmptyAddress Then
                
                    DeleteCustomerAddress strCustomerNumber, _
                                          strCustomerVersionNumber, _
                                          strAddressSeq, _
                                          strAddressGuid, _
                                          objADOAssist, objSQLAssist
                    
'                        objDOAssist.DeleteAllEx objXmlAddressNode, objXmlAddressSchema
'                        objDOAssist.DeleteAllEx objXmlCustomerAddressNode, objXmlCustomerAddressSchema
                
                Else
                    'SG 30/05/02 SYS4767 START
                    ' DRC 28/05/02 MSMS0065
                    ' check for existing customeraddress record
                    '
                    strTemp = "CUSTOMERNUMBER = '" & strCustomerNumber & "'" & _
                  " AND CUSTOMERVERSIONNUMBER = '" & strCustomerVersionNumber & "'" & _
                  " AND CUSTOMERADDRESSSEQUENCENUMBER = '" & strAddressSeq & "'"

                    Call objIADOAssist.GetValueFromTable("CUSTOMERADDRESS", strTemp, _
                                "ADDRESSGUID", strExistingAddressGUID, blnExistingCustomerAddress)
                
                    'Ensure that the existing 1:1 relationship between Customer and address is maintained
                    If (blnExistingCustomerAddress And (strExistingAddressGUID <> strAddressGuid)) Then
                        If Len(strExistingAddressGUID) = 0 Then
                          strAddressGuid = CreateGUID()
                          objXmlAssist.SetMandatoryChildText objXmlAddressNode, "ADDRESSGUID", strAddressGuid
                          vobjDOAssist.CreateEx objXmlAddressNode, objXmlAddressSchema
                        Else
                          strAddressGuid = strExistingAddressGUID
                        End If
                        objXmlAssist.SetMandatoryChildText objXmlCustomerAddressNode, "ADDRESSGUID", strAddressGuid
                    End If

                    If blnExistingCustomerAddress Then
                        objXmlAssist.SetMandatoryChildText objXmlAddressNode, "ADDRESSGUID", strAddressGuid
                        vobjDOAssist.UpdateEx objXmlAddressNode, objXmlAddressSchema
                        
                        ' SAB - 21/04/2006 - EP399 - Begin
                        Set xmlLastAmendedDate = objXmlCustomerAddressNode.selectSingleNode("LASTAMENDEDDATE")
                        
                        If xmlLastAmendedDate Is Nothing Then
                            Set xmlLastAmendedDate = objXmlCustomerAddressNode.ownerDocument.createElement("LASTAMENDEDDATE")
                            xmlLastAmendedDate.Text = Format(Now, "dd/mm/yyyy hh:mm:ss")
                            objXmlCustomerAddressNode.appendChild xmlLastAmendedDate
                        End If
                        ' SAB - 21/04/2006 - EP399 - End
                        
                        vobjDOAssist.UpdateEx objXmlCustomerAddressNode, objXmlCustomerAddressSchema
                    Else
                        strAddressGuid = CreateGUID()
                        objXmlAssist.SetMandatoryChildText objXmlAddressNode, "ADDRESSGUID", strAddressGuid
                        vobjDOAssist.CreateEx objXmlAddressNode, objXmlAddressSchema
                        objXmlAssist.SetMandatoryChildText objXmlCustomerAddressNode, "ADDRESSGUID", strAddressGuid
                        vobjDOAssist.CreateEx objXmlCustomerAddressNode, objXmlCustomerAddressSchema
                    End If
                    'SG 30/05/02 SYS4767 END
                End If
            
            Else
            
                If Not blnEmptyAddress Then
            
                    strAddressGuid = CreateGUID()
                
                    objXmlAssist.SetMandatoryChildText objXmlCustomerAddressNode, "CUSTOMERADDRESSSEQUENCENUMBER", Format(intNextSeq)
                    intNextSeq = intNextSeq + 1
                
                    objXmlAssist.SetMandatoryChildText objXmlCustomerAddressNode, "ADDRESSGUID", strAddressGuid
                    objXmlAssist.SetMandatoryChildText objXmlAddressNode, "ADDRESSGUID", strAddressGuid
                    
                    vobjDOAssist.CreateEx objXmlAddressNode, objXmlAddressSchema
                    vobjDOAssist.CreateEx objXmlCustomerAddressNode, objXmlCustomerAddressSchema
            
                End If
                
            End If
    
        Next
        
    End If

    ' clean up
    Set objSQLAssist = Nothing
    Set objADOAssist = Nothing
    Set objXmlAssist = Nothing
    Set objXmlCustomerAddressSchema = Nothing
    Set objXmlAddressSchema = Nothing
    Set objXmlCustomerAddressNode = Nothing
    Set objXmlAddressNode = Nothing
    Set objXmlNodeList = Nothing
    Set objXmlNode = Nothing
    Set objAddressClassDef = Nothing
    Set objAddressDO = Nothing
    Set objContext = Nothing
    Set xmlLastAmendedDate = Nothing ' EP399
    
    Exit Sub

WriteCustomerAddressesVbErr:

    ' clean up
    Set objSQLAssist = Nothing
    Set objADOAssist = Nothing
    Set objXmlAssist = Nothing
    Set objXmlCustomerAddressSchema = Nothing
    Set objXmlAddressSchema = Nothing
    Set objXmlCustomerAddressNode = Nothing
    Set objXmlAddressNode = Nothing
    Set objXmlNodeList = Nothing
    Set objXmlNode = Nothing
    Set objAddressClassDef = Nothing
    Set objAddressDO = Nothing
    Set objContext = Nothing
    Set xmlLastAmendedDate = Nothing ' EP399

    ' re-raise the error back up the calling chain
    Debug.Print Err.Number & ": " & Err.Description
    Err.Raise Err.Number, Err.Source, Err.Description, Err.HelpFile, Err.HelpContext

End Sub

Private Sub DeleteCustomerAddress(ByVal vstrCustomerNumber As String, _
                                  ByVal vstrCustomerVersionNumber As String, _
                                  ByVal vstrAddressSeq As String, _
                                  ByVal vstrAddressGuid As String, _
                                  ByVal vobjADOAssist As ADOAssist, _
                                  ByVal vobjSQLAssist As SQLAssist)
    
    Dim strSQL As String, strWhere As String
    
    ' delete CUSTOMERADDRESS
    strWhere = "CUSTOMERNUMBER=" & vobjSQLAssist.FormatString(vstrCustomerNumber)
    strWhere = strWhere & " AND CUSTOMERVERSIONNUMBER=" & vstrCustomerVersionNumber
    strWhere = strWhere & " AND CUSTOMERADDRESSSEQUENCENUMBER=" & vstrAddressSeq
    
    strSQL = "DELETE FROM CUSTOMERADDRESS WHERE " & strWhere
    
    vobjADOAssist.executeSQLCommand strSQL
    
    ' delete ADDRESS
    strWhere = "ADDRESSGUID=" & vobjSQLAssist.FormatGuid(vstrAddressGuid)
    
    strSQL = "DELETE FROM ADDRESS WHERE " & strWhere
    
    vobjADOAssist.executeSQLCommand strSQL

End Sub


Private Function GetHighestAddressSeqNo(ByVal vstrCustomerNumber As String, _
                                        ByVal vstrCustomerVersionNumber As String, _
                                        ByVal vobjADOAssist As ADOAssist, _
                                        ByVal vobjSQLAssist As SQLAssist) _
                                        As Integer
    
On Error GoTo GetHighestAddressSeqNoErr

    Dim cmd As New ADODB.Command
    Dim rst As ADODB.Recordset
    
    Dim strSQL As String
    'DM     15/10/2002  BMIDS00935 Added WITH (NOLOCK) to this SQL string
    strSQL = _
        "select max(CUSTOMERADDRESSSEQUENCENUMBER) " & _
        " from CUSTOMERADDRESS WITH (NOLOCK)" & _
        " where CUSTOMERNUMBER = " & vobjSQLAssist.FormatString(vstrCustomerNumber) & _
        " and CUSTOMERVERSIONNUMBER = " & vstrCustomerVersionNumber
    
    cmd.CommandText = strSQL
    
    cmd.ActiveConnection = vobjADOAssist.GetConnStr
    
    Set rst = cmd.Execute(1, , adCmdText)
    
    ' rst!FieldName does not work where field value is null
    If IsNull(rst.Fields(0).Value) Then
        ' sequence number not yet set on this field
        GetHighestAddressSeqNo = 0
    Else
        GetHighestAddressSeqNo = rst.Fields(0).Value
    End If
    rst.Close
    cmd.ActiveConnection.Close
    
    Set rst = Nothing
    Set cmd = Nothing

GetHighestAddressSeqNoErr:
    ' Close the recordset if still open
        If Not rst Is Nothing Then
                If rst.State = adStateOpen Then
                        rst.Close
                End If
        End If

        Set rst = Nothing
        Set cmd = Nothing

End Function

Private Sub WriteCustomerTelephones(ByVal vstrCustomerNumber As String, _
                                    ByVal vstrCustomerVersionNumber As String, _
                                    ByVal vobjCustomerTelephoneList As IXMLDOMNode)

On Error GoTo WriteCustomerTelephonesVbErr
    
    Dim objContext As ObjectContext
    Set objContext = GetObjectContext()

    Dim objXmlNodeList As IXMLDOMNodeList
    Dim objXmlNode As IXMLDOMNode
    Dim xmlTempResponseNode As IXMLDOMNode
    Dim xmlResponseNode As IXMLDOMElement
    Dim xmlDoc As New FreeThreadedDOMDocument40
    xmlDoc.validateOnParse = False
    xmlDoc.setProperty "NewParser", True
    
    Dim objXmlAssist As New XMLAssist
    
    Dim objXmlSchemaDoc As FreeThreadedDOMDocument40
    
    Dim objClassDefs As IomCustClassDef
    Set objClassDefs = New omCustClassDef
    
    Dim objADOAssist As ADOAssist
    Dim objSQLAssist As SQLAssist
    Dim objErrAssist As ErrAssist
    Dim objDOAssist  As DOAssist
    Dim objICustomerBO As ICustomerBO
    Set objSQLAssist = New SQLAssist

    If objContext Is Nothing Then

        Set objADOAssist = New ADOAssist
        Set objDOAssist = New DOAssist
        Set objICustomerBO = New CustomerBO
    Else

        Set objADOAssist = _
            objContext.CreateInstance(gstrBASE_COMPONENT & ".ADOAssist")
        Set objDOAssist = _
            objContext.CreateInstance(gstrBASE_COMPONENT & ".DOAssist")
        Set objICustomerBO = _
            objContext.CreateInstance(gstrCUSTOMER_COMPONENT & ".CustomerBO")
    End If
            
    Dim IntCount As Integer, intSeq As Integer
    
    Set objErrAssist = New ErrAssist
    DeleteCustomerTelephones vstrCustomerNumber, _
                             vstrCustomerVersionNumber, _
                             objADOAssist, objSQLAssist
    
    Set objXmlSchemaDoc = objClassDefs.LoadCustomerTelephoneData
    Set objXmlNodeList = vobjCustomerTelephoneList.selectNodes("CUSTOMERTELEPHONENUMBER")
    
    If objXmlNodeList.length > 0 Then
        
        intSeq = 1
        
        For IntCount = 0 To (objXmlNodeList.length - 1)
            
            Set objXmlNode = objXmlNodeList.Item(IntCount)
            
            If Not IsEmptyTelephoneNode(objXmlNode) Then
                
                objXmlAssist.SetMandatoryChildText objXmlNode, "TELEPHONESEQUENCENUMBER", Format(intSeq)
                intSeq = intSeq + 1
                ' Validate the Telphone Numbers (SR / AQR : SYS0085)
                Set xmlTempResponseNode = objICustomerBO.ValidateCustomerTelephone(objXmlNode)
                Set xmlResponseNode = xmlDoc.createElement("RESPONSE")
                objErrAssist.CheckResponse xmlTempResponseNode, xmlResponseNode, True
                objDOAssist.CreateEx objXmlNode, objXmlSchemaDoc
            End If
        Next
        
    End If
    
    Set objClassDefs = Nothing
    Set objADOAssist = Nothing
    Set objICustomerBO = Nothing
    Set objSQLAssist = Nothing
    Set objXmlSchemaDoc = Nothing
    Set objXmlNodeList = Nothing
    Set objXmlNode = Nothing
    Set xmlResponseNode = Nothing
    Set xmlTempResponseNode = Nothing
    Set xmlDoc = Nothing
    Set objXmlAssist = Nothing
    Set objContext = Nothing
    
    Exit Sub
    
WriteCustomerTelephonesVbErr:
    
    Set objClassDefs = Nothing
    Set objADOAssist = Nothing
    Set objICustomerBO = Nothing
    Set objSQLAssist = Nothing
    Set objXmlSchemaDoc = Nothing
    Set objXmlNodeList = Nothing
    Set objXmlNode = Nothing
    Set xmlResponseNode = Nothing
    Set xmlTempResponseNode = Nothing
    Set xmlDoc = Nothing
    Set objXmlAssist = Nothing
    Set objContext = Nothing

    ' re-raise the error back up the calling chain
    Debug.Print Err.Number & ": " & Err.Description
    Err.Raise Err.Number, Err.Source, Err.Description, Err.HelpFile, Err.HelpContext

End Sub

' BMIDS864 New function to save verification
Private Sub WriteCustomerVerifications(ByVal vstrCustomerNumber As String, _
                                       ByVal vstrCustomerVersionNumber As String, _
                                       ByVal vobjCustomerVerificationList As IXMLDOMNode)

On Error GoTo WriteCustomerVerificationsVbErr
    
    Dim objContext As ObjectContext
    Set objContext = GetObjectContext()

    Dim objXmlNodeList As IXMLDOMNodeList
    Dim objXmlNode As IXMLDOMNode
    Dim xmlDoc As New FreeThreadedDOMDocument40
    xmlDoc.validateOnParse = False
    xmlDoc.setProperty "NewParser", True
    
    Dim objXmlAssist As New XMLAssist
    
    Dim objXmlSchemaDoc As FreeThreadedDOMDocument40
    
    Dim objClassDefs As IomCustClassDef
    Set objClassDefs = New omCustClassDef
    
    Dim objADOAssist As ADOAssist
    Dim objSQLAssist As SQLAssist
    Dim objErrAssist As ErrAssist
    Dim objDOAssist  As DOAssist
    Dim objICustomerBO As ICustomerBO
    Set objSQLAssist = New SQLAssist

    If objContext Is Nothing Then

        Set objADOAssist = New ADOAssist
        Set objDOAssist = New DOAssist
        Set objICustomerBO = New CustomerBO
    Else

        Set objADOAssist = _
            objContext.CreateInstance(gstrBASE_COMPONENT & ".ADOAssist")
        Set objDOAssist = _
            objContext.CreateInstance(gstrBASE_COMPONENT & ".DOAssist")
        Set objICustomerBO = _
            objContext.CreateInstance(gstrCUSTOMER_COMPONENT & ".CustomerBO")
    End If
            
    Dim IntCount As Integer, intSeq As Integer
    
    Set objErrAssist = New ErrAssist
    DeleteCustomerVerifications vstrCustomerNumber, _
                                vstrCustomerVersionNumber, _
                                objADOAssist, objSQLAssist
    
    Set objXmlSchemaDoc = objClassDefs.LoadVerificationData
    Set objXmlNodeList = vobjCustomerVerificationList.selectNodes("VERIFICATION")
    
    If objXmlNodeList.length > 0 Then
        
        intSeq = 1
        
        For IntCount = 0 To (objXmlNodeList.length - 1)
            
            Set objXmlNode = objXmlNodeList.Item(IntCount)
            
            If Not IsEmptyVerificationNode(objXmlNode) Then
                
                objXmlAssist.SetMandatoryChildText objXmlNode, "VERIFICATIONSEQUENCENUMBER", Format(intSeq)
                intSeq = intSeq + 1
                
                'Set the Verification Type = 1 as BMIDS only return Electronic ID.
                objXmlAssist.SetMandatoryChildText objXmlNode, "VERIFICATIONTYPE", 1
                
                objDOAssist.CreateEx objXmlNode, objXmlSchemaDoc
            End If
        Next
        
    End If
    
    Set objClassDefs = Nothing
    Set objADOAssist = Nothing
    Set objICustomerBO = Nothing
    Set objSQLAssist = Nothing
    Set objXmlSchemaDoc = Nothing
    Set objXmlNodeList = Nothing
    Set objXmlNode = Nothing
    Set xmlDoc = Nothing
    Set objXmlAssist = Nothing
    Set objContext = Nothing
    
    Exit Sub
    
WriteCustomerVerificationsVbErr:
    
    Set objClassDefs = Nothing
    Set objADOAssist = Nothing
    Set objICustomerBO = Nothing
    Set objSQLAssist = Nothing
    Set objXmlSchemaDoc = Nothing
    Set objXmlNodeList = Nothing
    Set objXmlNode = Nothing
    Set xmlDoc = Nothing
    Set objXmlAssist = Nothing
    Set objContext = Nothing

    ' re-raise the error back up the calling chain
    Debug.Print Err.Number & ": " & Err.Description
    Err.Raise Err.Number, Err.Source, Err.Description, Err.HelpFile, Err.HelpContext

End Sub

Private Sub DeleteCustomerVerifications(ByVal vstrCustomerNumber As String, _
                                        ByVal vstrCustomerVersionNumber As String, _
                                        ByVal vobjADOAssist As ADOAssist, _
                                        ByVal vobjSQLAssist As SQLAssist)
    
On Error Resume Next
    
    Dim strSQL As String, strWhere As String
    
    strWhere = "CUSTOMERNUMBER=" & vobjSQLAssist.FormatString(vstrCustomerNumber)
    strWhere = strWhere & " AND CUSTOMERVERSIONNUMBER=" & vstrCustomerVersionNumber
    
    strSQL = "DELETE FROM VERIFICATION WHERE " & strWhere
    
    vobjADOAssist.executeSQLCommand strSQL
    
    ' PSC 16/01/2006 MAR1050 - Start
    If errGetOmigaErrorNumber(Err.Number) = oeNoRowsAffected Then
        Err.Clear
    End If
    ' PSC 16/01/2006 MAR1050 - End
    
End Sub

Private Function IsEmptyVerificationNode(ByVal vobjXmlVerification As IXMLDOMNode) As Boolean
        
    If Not vobjXmlVerification.selectSingleNode("VERIFICATIONTYPE") Is Nothing Then
        If Len(vobjXmlVerification.selectSingleNode("VERIFICATIONTYPE").Text) > 0 Then
            IsEmptyVerificationNode = False
            Exit Function
        End If
    End If
    
    If Not vobjXmlVerification.selectSingleNode("IDENTIFICATIONTYPE") Is Nothing Then
        If Len(vobjXmlVerification.selectSingleNode("IDENTIFICATIONTYPE").Text) > 0 Then
            IsEmptyVerificationNode = False
            Exit Function
        End If
    End If
    
    If Not vobjXmlVerification.selectSingleNode("OTHERIDENTIFICATION") Is Nothing Then
        If Len(vobjXmlVerification.selectSingleNode("OTHERIDENTIFICATION").Text) > 0 Then
            IsEmptyVerificationNode = False
            Exit Function
        End If
    End If
    
 
    If Not vobjXmlVerification.selectSingleNode("REFERENCE") Is Nothing Then
        If Len(vobjXmlVerification.selectSingleNode("REFERENCE").Text) > 0 Then
            IsEmptyVerificationNode = False
            Exit Function
        End If
    End If

    If Not vobjXmlVerification.selectSingleNode("ISSUER") Is Nothing Then
        If Len(vobjXmlVerification.selectSingleNode("ISSUER").Text) > 0 Then
            IsEmptyVerificationNode = False
            Exit Function
        End If
    End If

    If Not vobjXmlVerification.selectSingleNode("VERIFICATIONDATE") Is Nothing Then
        If Len(vobjXmlVerification.selectSingleNode("VERIFICATIONDATE").Text) > 0 Then
            IsEmptyVerificationNode = False
            Exit Function
        End If
    End If

    IsEmptyVerificationNode = True

End Function

' *** End

Private Sub DeleteCustomerTelephones(ByVal vstrCustomerNumber As String, _
                                     ByVal vstrCustomerVersionNumber As String, _
                                     ByVal vobjADOAssist As ADOAssist, _
                                     ByVal vobjSQLAssist As SQLAssist)
    
On Error Resume Next
    
    Dim strSQL As String, strWhere As String
    
    ' delete CUSTOMERTELEPHONENUMBER
    strWhere = "CUSTOMERNUMBER=" & vobjSQLAssist.FormatString(vstrCustomerNumber)
    strWhere = strWhere & " AND CUSTOMERVERSIONNUMBER=" & vstrCustomerVersionNumber
    
    strSQL = "DELETE FROM CUSTOMERTELEPHONENUMBER WHERE " & strWhere
    
    vobjADOAssist.executeSQLCommand strSQL
    
    ' PSC 16/01/2006 MAR1050 - Start
    If errGetOmigaErrorNumber(Err.Number) = oeNoRowsAffected Then
        Err.Clear
    End If
    ' PSC 16/01/2006 MAR1050 - End
    
End Sub

Private Function IsEmptyTelephoneNode(ByVal vobjXmlTelephone As IXMLDOMNode) As Boolean
        
    If Not vobjXmlTelephone.selectSingleNode("TELEPHONENUMBER") Is Nothing Then
        If Len(vobjXmlTelephone.selectSingleNode("TELEPHONENUMBER").Text) > 0 Then
            IsEmptyTelephoneNode = False
            Exit Function
        End If
    End If
    
    If Not vobjXmlTelephone.selectSingleNode("USAGE") Is Nothing Then
        If Len(vobjXmlTelephone.selectSingleNode("USAGE").Text) > 0 Then
            IsEmptyTelephoneNode = False
            Exit Function
        End If
    End If
    
    If Not vobjXmlTelephone.selectSingleNode("CONTACTTIME") Is Nothing Then
        If Len(vobjXmlTelephone.selectSingleNode("CONTACTTIME").Text) > 0 Then
            IsEmptyTelephoneNode = False
            Exit Function
        End If
    End If
    
    'JR - Omiplus24
    If Not vobjXmlTelephone.selectSingleNode("COUNTRYCODE") Is Nothing Then
        If Len(vobjXmlTelephone.selectSingleNode("COUNTRYCODE").Text) > 0 Then
            IsEmptyTelephoneNode = False
            Exit Function
        End If
    End If
        
    If Not vobjXmlTelephone.selectSingleNode("AREACODE") Is Nothing Then
        If Len(vobjXmlTelephone.selectSingleNode("AREACODE").Text) > 0 Then
            IsEmptyTelephoneNode = False
            Exit Function
        End If
    End If
    'JR End

    IsEmptyTelephoneNode = True

End Function

Private Sub WriteAreasOfInterest(ByVal vstrCustomerNumber As String, _
                                 ByVal vstrCustomerVersionNumber As String, _
                                 ByVal vobjAreasOfInterestList As IXMLDOMNode)

On Error GoTo WriteAreasOfInterestVbErr
    
    Dim objContext As ObjectContext
    Set objContext = GetObjectContext()

    Dim objXmlNodeList As IXMLDOMNodeList
    Dim objXmlNode As IXMLDOMNode
    
    Dim objXmlAssist As New XMLAssist
    
    Dim objXmlSchemaDoc As FreeThreadedDOMDocument40
    
    Dim objClassDefs As IomCustClassDef
    Set objClassDefs = New omCustClassDef
    
    Dim objADOAssist As ADOAssist
    Dim objSQLAssist As SQLAssist
    Dim objDOAssist  As DOAssist
    Set objSQLAssist = New SQLAssist

    If objContext Is Nothing Then
        Set objADOAssist = New ADOAssist
        Set objDOAssist = New DOAssist
    Else
        Set objADOAssist = _
            objContext.CreateInstance(gstrBASE_COMPONENT & ".ADOAssist")
        Set objDOAssist = _
            objContext.CreateInstance(gstrBASE_COMPONENT & ".DOAssist")
    End If
    
    Dim IntCount As Integer
    
    DeleteAreasOfInterest vstrCustomerNumber, _
                          vstrCustomerVersionNumber, _
                          objADOAssist, objSQLAssist
    
    Set objXmlSchemaDoc = objClassDefs.LoadAreasOfInterestData
    Set objXmlNodeList = vobjAreasOfInterestList.selectNodes("AREASOFINTEREST")
    
    If objXmlNodeList.length > 0 Then
        
        For IntCount = 0 To (objXmlNodeList.length - 1)
            
            Set objXmlNode = objXmlNodeList.Item(IntCount)
            
            If Not IsEmptyAreasOfInterest(objXmlNode) Then
                
                objDOAssist.CreateEx objXmlNode, objXmlSchemaDoc
                
            End If
        
        Next
        
    End If
    
    Set objClassDefs = Nothing
    Set objADOAssist = Nothing
    Set objSQLAssist = Nothing
    Set objXmlSchemaDoc = Nothing
    Set objXmlNodeList = Nothing
    Set objXmlNode = Nothing
    Set objXmlAssist = Nothing
    Set objContext = Nothing
    
    Exit Sub
    
WriteAreasOfInterestVbErr:
    
    Set objClassDefs = Nothing
    Set objADOAssist = Nothing
    Set objSQLAssist = Nothing
    Set objXmlSchemaDoc = Nothing
    Set objXmlNodeList = Nothing
    Set objXmlNode = Nothing
    Set objXmlAssist = Nothing
    Set objContext = Nothing

    ' re-raise the error back up the calling chain
    Debug.Print Err.Number & ": " & Err.Description
    Err.Raise Err.Number, Err.Source, Err.Description, Err.HelpFile, Err.HelpContext

End Sub

Private Function IsEmptyAreasOfInterest(ByVal vobjXmlAreaOfInterest As IXMLDOMNode) As Boolean
    
    If Not vobjXmlAreaOfInterest.selectSingleNode("INTERESTAREA") Is Nothing Then
        If Len(vobjXmlAreaOfInterest.selectSingleNode("INTERESTAREA").Text) > 0 Then
            IsEmptyAreasOfInterest = False
            Exit Function
        End If
    End If

    IsEmptyAreasOfInterest = True

End Function

Private Sub DeleteAreasOfInterest(ByVal vstrCustomerNumber As String, _
                                  ByVal vstrCustomerVersionNumber As String, _
                                  ByVal vobjADOAssist As ADOAssist, _
                                  ByVal vobjSQLAssist As SQLAssist)
    
On Error Resume Next
    
    Dim strSQL As String, strWhere As String
    
    strWhere = "CUSTOMERNUMBER=" & vobjSQLAssist.FormatString(vstrCustomerNumber)
    strWhere = strWhere & " AND CUSTOMERVERSIONNUMBER=" & vstrCustomerVersionNumber
    
    strSQL = "DELETE FROM AREASOFINTEREST WHERE " & strWhere
    
    vobjADOAssist.executeSQLCommand strSQL
    
    ' PSC 16/01/2006 MAR1050 - Start
    If errGetOmigaErrorNumber(Err.Number) = oeNoRowsAffected Then
        Err.Clear
    End If
    ' PSC 16/01/2006 MAR1050 - End
    
End Sub

Private Sub ICustomerDO_CustomerDetailsDerivedData(ByVal vobjXmlCustomerNode As IXMLDOMNode)
' header ----------------------------------------------------------------------------------
' description:
'   XML elements must be created for any derived values as specified.
'   Add any derived values to XML. E.g. data type 'double' fields will
' pass:
'   vstrXMLData         base XML data stream
'                       as:
'                           <tablename>
'                               <element1>element1 value</element1>
'                               <elementn>elementn value</elementn>
' return:
'   AddDerivedData      base XML data stream plus any derived values
' Raise Errors:
'------------------------------------------------------------------------------------------
    
    Dim objXmlNode As IXMLDOMNode
    Dim objXmlElem As IXMLDOMElement
    
    Dim intAge As Integer
                
    Set objXmlNode = vobjXmlCustomerNode.selectSingleNode("CUSTOMERVERSION/DATEOFBIRTH")
    
    If Not objXmlNode Is Nothing Then
    
        intAge = GetAge(objXmlNode.Text)
        
        Set objXmlElem = vobjXmlCustomerNode.ownerDocument.createElement("AGE")
        objXmlElem.Text = Format(intAge)
        objXmlNode.parentNode.appendChild objXmlElem
    
    End If

End Sub

Private Function GetAge(ByVal vstrDateOfBirth As String) As Integer
' header ----------------------------------------------------------------------------------
' description:          calculate Age from date of birth
' pass:
'       strDateOfBirth  date of birth in format dd/mm/yyyy
' return:
'       GetAge  integer representation of age as calculated from date of birth
' Raise Errors:
'------------------------------------------------------------------------------------------
On Error GoTo GetAgeVbErr

    Const strFunctionName As String = "GetAge"
    
    Dim objErrAssist As New ErrAssist
    
    If (Len(vstrDateOfBirth) < Len("dd/mm/yyyy")) Or _
       (Not IsDate(vstrDateOfBirth)) Then
        objErrAssist.ThrowError TypeName(Me), strFunctionName, oeInvalidDateTimeFormat, _
                        " Expected Date Of Birth in format DD/MM/YYYY, but received: " & vstrDateOfBirth
    End If
    
    Dim dtmDoB As Date
    Dim lngDoB As Long
    Dim lngNow As Long
    
    dtmDoB = vstrDateOfBirth
    
    lngDoB = (Year(dtmDoB) * 10000) + (Month(dtmDoB) * 100) + Day(dtmDoB)
    lngNow = (Year(Now) * 10000) + (Month(Now) * 100) + Day(Now)
    
    GetAge = Fix((lngNow - lngDoB) / 10000)
    
    Set objErrAssist = Nothing
    
    Exit Function
    
GetAgeVbErr:
    
    Set objErrAssist = Nothing

    '   re-raise error for business object to interpret as appropriate
    Err.Raise Err.Number, Err.Source, Err.Description

End Function

Private Sub AddSalutationToRequest(ByVal robjXmlCustomerVersionNode As IXMLDOMNode)
' header ----------------------------------------------------------------------------------
' description: If the Salutation does not exist then add it as a child element to the
' parent see derived Data.xls for specific information concerning how to derive the
' salutation
'
' pass: robjXmlCustomerElem     an XML element
'
' return:   the amended XML element
'
' Raise Errors:     n/a
'------------------------------------------------------------------------------------------
    
On Error GoTo AddSalutationToRequestVbErr
    
    Dim objXmlElem As IXMLDOMElement
    Dim objXmlNode As IXMLDOMNode
        
    Dim objComboDO As New ComboDO

    Dim objXmlAssist As New XMLAssist
    
    Dim strSalutation As String
    Dim strTitleId As String
    Dim strSurname As String
    Dim strTitleText As String
    Dim strGroupName As String
    Dim blnTagExists As Boolean
    '
    ' if the salutation tag did not exist then we need to create one
    '
    
    Set objXmlNode = robjXmlCustomerVersionNode.selectSingleNode("CORRESPONDENCESALUTATION")
    
    If objXmlNode Is Nothing Then
        Set objXmlElem = _
            robjXmlCustomerVersionNode.ownerDocument.createElement("CORRESPONDENCESALUTATION")
        Set objXmlNode = robjXmlCustomerVersionNode.appendChild(objXmlElem)
   
    End If
    
    ' create the salutation to be = Title + " " + Surname
    ' first check for an other title and use it if it exists as the start of the salutation
    If objXmlAssist.GetElementTextBln(robjXmlCustomerVersionNode, "TITLEOTHER", strTitleText) Then
        strSalutation = strTitleText
    End If
    ' otherwise use the combo title
    If Len(strTitleText) = 0 Then
        If objXmlAssist.GetElementTextBln(robjXmlCustomerVersionNode, "TITLE", strTitleId) Then
            strTitleText = objComboDO.GetComboText("Title", strTitleId)
            strSalutation = strTitleText
        End If
    End If
    If objXmlAssist.GetElementTextBln(robjXmlCustomerVersionNode, "SURNAME", strSurname) Then
        
        If Len(strSalutation) > 0 Then
            strSalutation = strSalutation & " "
        End If
        strSalutation = strSalutation & strSurname
    End If
    objXmlNode.Text = strSalutation

    
    Set objXmlAssist = Nothing
    Set objXmlElem = Nothing
    Set objXmlNode = Nothing
    Set objComboDO = Nothing
    
    Exit Sub

AddSalutationToRequestVbErr:

    Set objXmlAssist = Nothing
    Set objXmlElem = Nothing
    Set objXmlNode = Nothing
    Set objComboDO = Nothing

    ' re-raise the error back up the calling chain
    Err.Raise Err.Number, Err.Source, Err.Description, Err.HelpFile, Err.HelpContext

End Sub

Private Function LoadClassDefForGetSpecifiedData(ByVal strTable As String) As FreeThreadedDOMDocument40
'-----------------------------------------------------------------------------------------
'Description : Loads the relevant class definition based on the table passed
'Pass        : strTable : table Name
'Return      : DOMDocument
'-----------------------------------------------------------------------------------------

On Error GoTo LoadClassDefForGetSpecifiedDataVbErr

    Const strFunctionName As String = "LoadClassDefForGetSpecifiedData"

    Dim objXmlAssist As New XMLAssist
    Dim objErrAssist As New ErrAssist
    Dim objICustClassDef As IomCustClassDef
    Set objICustClassDef = New omCustClassDef
    
    Select Case UCase(strTable)
        Case "CUSTOMER"
            Set LoadClassDefForGetSpecifiedData = objICustClassDef.LoadCustomerData
        Case "CUSTOMERVERSION"
            Set LoadClassDefForGetSpecifiedData = objICustClassDef.LoadCustomerVersionData
        Case Else
            objErrAssist.ThrowError TypeName(Me), strFunctionName, 507, _
                            "Unable to identify the class definition"
    End Select
        
    Set objXmlAssist = Nothing
    Set objErrAssist = Nothing
    Set objICustClassDef = Nothing
    
    Exit Function
LoadClassDefForGetSpecifiedDataVbErr:
    
    Set objXmlAssist = Nothing
    Set objICustClassDef = Nothing

    objErrAssist.AddToErrSource strFunctionName

    If objErrAssist.IsSystemError = True Then
        objErrAssist.LogError TypeName(Me), strFunctionName, Err.Number, Err.Description
    End If

End Function

Private Sub ObjectControl_Activate()

    Set m_objContext = GetObjectContext()
    Set m_objErrAssist = New ErrAssist
    Set m_objXmlAssist = New XMLAssist
    
    If m_objContext Is Nothing Then
        Set m_objIDOAssist = New DOAssist
    Else
        Set m_objIDOAssist = m_objContext.CreateInstance(gstrBASE_COMPONENT & ".DOAssist")
    End If

End Sub

Private Function ObjectControl_CanBePooled() As Boolean

    ObjectControl_CanBePooled = False

End Function

Private Sub ObjectControl_Deactivate()

    Set m_objContext = Nothing
    Set m_objErrAssist = Nothing
    Set m_objXmlAssist = Nothing
    Set m_objIDOAssist = Nothing

End Sub

' BMIDS00005 CRWP2 BM065 ICustomerDO_FindContactsPerApplication
Private Function ICustomerDO_FindContactsPerApplication(ByVal vxmlTableElement As IXMLDOMElement) As IXMLDOMNode
' header ----------------------------------------------------------------------------------
' description:  Finds Contacts for an application by querying a view
' pass:         vxmlTableElement  xml element containing the application number to query on
' return:       xml node containing retrieved contacts
'------------------------------------------------------------------------------------------
On Error GoTo FindContactsPerApplicationVbErr

    Const strFunctionName As String = "ICustomerDO_FindContactsPerApplication"
    
                           
    Dim xmlClassDefDoc  As FreeThreadedDOMDocument40
    Dim objIClassDef    As IomCustClassDef
    Dim xmlResults      As IXMLDOMNode
    Dim xmlDoc          As FreeThreadedDOMDocument40
    Dim xmlElement      As IXMLDOMElement
    
    Set xmlDoc = New FreeThreadedDOMDocument40
    xmlDoc.validateOnParse = False
    xmlDoc.setProperty "NewParser", True
    
    Set xmlElement = xmlDoc.createElement("FINDCONTACTSPERAPPLICATION")
    xmlElement.appendChild vxmlTableElement.selectSingleNode("CONTACTLOG/APPLICATIONNUMBER").cloneNode(True)
    
    Set objIClassDef = New omCustClassDef
    Set xmlClassDefDoc = objIClassDef.LoadApplicationContactData()
    
    Set xmlResults = m_objIDOAssist.FindListMultiple(xmlElement, xmlClassDefDoc)
    
    If Not xmlResults Is Nothing Then   'BMIDS00325 If nothing is found then skip this step
        xmlChangeNodeName xmlResults, "FINDCONTACTSPERAPPLICATIONLIST", "CONTACTLOG"
        For Each xmlElement In xmlResults.selectNodes("FINDCONTACTSPERAPPLICATION")
            xmlChangeNodeName xmlElement, "FINDCONTACTSPERAPPLICATION", "CONTACT"
        Next
        Set ICustomerDO_FindContactsPerApplication = xmlResults
    'BMIDS00325
    Else
        Set ICustomerDO_FindContactsPerApplication = Nothing
    End If
    'BMIDS00325 End
          
    m_objContext.SetComplete
        
FindContactsPerApplicationVbErr:
    
    'Record not found errors can be ignored here as they are non fatal
    If m_objErrAssist.GetOmigaErrorNumber(Err.Number) = oeRecordNotFound Then
        Resume Next
    End If
    
    
    Set xmlClassDefDoc = Nothing
    Set objIClassDef = Nothing
    Set xmlResults = Nothing
    Set xmlDoc = Nothing
    Set xmlElement = Nothing
    If Err.Number = 0 Then
        Exit Function
    End If
    
    m_objErrAssist.AddToErrSource (strFunctionName)
    
    If m_objErrAssist.IsSystemError = True Then
        m_objErrAssist.LogError TypeName(Me), strFunctionName, Err.Number, Err.Description
    End If
    
    If m_objErrAssist.IsSystemError Then
        m_objContext.SetAbort
    Else
        m_objContext.SetComplete
    End If
        
    '   re-raise error for business object to interpret as appropriate
    Err.Raise Err.Number, Err.Source, Err.Description
End Function
'BMIDS00005 End

' BMIDS00026 IWP2 BM060 ICustomerDO_FindCustomersPerApplication
Private Function ICustomerDO_FindCustomersPerApplication(ByVal vxmlTableElement As IXMLDOMElement) As IXMLDOMNode
' header ----------------------------------------------------------------------------------
' description:  Finds details about all customers liked to an application
' pass:         vxmlTableElement  xml element containing the application number to query on
' return:       xml node containing retrieved customers
'------------------------------------------------------------------------------------------
On Error GoTo FindCustomersPerApplicationVbErr

    Const strFunctionName As String = "ICustomerDO_FindCustomersPerApplication"
    
                           
    Dim xmlClassDefDoc              As FreeThreadedDOMDocument40
    Dim xmlPhoneClassDef            As FreeThreadedDOMDocument40
    Dim xmlVerificationClassDef     As FreeThreadedDOMDocument40
    Dim xmlAddressClassDef          As FreeThreadedDOMDocument40  'BMIDS00205
    Dim xmlDoc                      As FreeThreadedDOMDocument40
    Dim objIClassDef                As IomCustClassDef
    Dim xmlResults                  As IXMLDOMNode
    Dim xmlSubResults               As IXMLDOMElement
    Dim xmlElement                  As IXMLDOMElement
    Dim xmlCustomerNumber           As IXMLDOMElement
    Dim xmlCustomerVersion          As IXMLDOMElement
    Dim xmlCustomer                 As IXMLDOMNode
    Dim xmlSubNode                  As IXMLDOMNode
    Dim strCustomerNumber           As String
    Dim strCustomerVersionNumber    As String
                
    Set xmlDoc = New FreeThreadedDOMDocument40
    xmlDoc.validateOnParse = False
    xmlDoc.setProperty "NewParser", True
    Set xmlElement = xmlDoc.createElement("FINDCUSTOMERSPERAPPLICATION")
    xmlElement.appendChild vxmlTableElement.selectSingleNode("CUSTOMERS/APPLICATIONNUMBER").cloneNode(True)
      
    Set objIClassDef = New omCustClassDef
    Set xmlClassDefDoc = objIClassDef.LoadApplicationCustomerData()

    Set xmlResults = m_objIDOAssist.FindListMultiple(xmlElement, xmlClassDefDoc)

    Set xmlPhoneClassDef = objIClassDef.LoadCustomerPhoneData()
    Set xmlVerificationClassDef = objIClassDef.LoadCustomerVerificationData()
    'BMIDS00205 15/07/2002 GHun
    Set xmlAddressClassDef = objIClassDef.LoadCRSCustomerAddressData()
    
    xmlChangeNodeName xmlResults, "FINDCUSTOMERSPERAPPLICATIONLIST", "CUSTOMERS"
    
    'Loop through each of the customers found and retrieve phone numbers and verification details
    For Each xmlCustomer In xmlResults.selectNodes("FINDCUSTOMERSPERAPPLICATION")
        xmlChangeNodeName xmlCustomer, "FINDCUSTOMERSPERAPPLICATION", "CUSTOMER"
        
        strCustomerNumber = xmlGetMandatoryNodeText(xmlCustomer, "OMIGACUSTOMERNUMBER")
        Set xmlCustomerNumber = xmlDoc.createElement("CUSTOMERNUMBER")
        xmlCustomerNumber.Text = strCustomerNumber
        
        strCustomerVersionNumber = xmlGetMandatoryNodeText(xmlCustomer, "CUSTOMERVERSIONNUMBER")
        Set xmlCustomerVersion = xmlDoc.createElement("CUSTOMERVERSIONNUMBER")
        xmlCustomerVersion.Text = strCustomerVersionNumber
        
        ' Find phone numbers for the current customer
        Set xmlElement = xmlDoc.createElement("FINDPHONENUMBERSPERCUSTOMER")
        
        xmlElement.appendChild xmlCustomerNumber.cloneNode(True)
        xmlElement.appendChild xmlCustomerVersion.cloneNode(True)
        
        Set xmlSubResults = m_objIDOAssist.FindListMultiple(xmlElement, xmlPhoneClassDef)

        If Not xmlSubResults Is Nothing Then
            For Each xmlSubNode In xmlSubResults.selectNodes("FINDPHONENUMBERSPERCUSTOMER")
                xmlChangeNodeName xmlSubNode, "FINDPHONENUMBERSPERCUSTOMER", "PHONENUMBER"
                xmlSubNode.removeChild xmlSubNode.selectSingleNode("CUSTOMERNUMBER")
                xmlSubNode.removeChild xmlSubNode.selectSingleNode("CUSTOMERVERSIONNUMBER")
                
                xmlCustomer.appendChild xmlSubNode.cloneNode(True)
            Next
        End If

        ' Find verification details for the current customer
        Set xmlElement = xmlDoc.createElement("FINDVERIFICATIONPERCUSTOMER")
               
        xmlElement.appendChild xmlCustomerNumber.cloneNode(True)
        xmlElement.appendChild xmlCustomerVersion.cloneNode(True)
        xmlElement.appendChild vxmlTableElement.selectSingleNode("CUSTOMERS/APPLICATIONNUMBER").cloneNode(True)
        
        Set xmlSubResults = Nothing
        Set xmlSubResults = m_objIDOAssist.FindListMultiple(xmlElement, xmlVerificationClassDef)

        If Not xmlSubResults Is Nothing Then
            For Each xmlSubNode In xmlSubResults.selectNodes("FINDVERIFICATIONPERCUSTOMER")
                xmlChangeNodeName xmlSubNode, "FINDVERIFICATIONPERCUSTOMER", "VERIFICATION"
                xmlSubNode.removeChild xmlSubNode.selectSingleNode("CUSTOMERNUMBER")
                xmlSubNode.removeChild xmlSubNode.selectSingleNode("CUSTOMERVERSIONNUMBER")
                xmlSubNode.removeChild xmlSubNode.selectSingleNode("APPLICATIONNUMBER")
                
                xmlCustomer.appendChild xmlSubNode.cloneNode(True)
            Next
        End If
        
        ' BMIDS00205 15/07/2002 GHun
        ' Find addresses for the current customer
        Set xmlElement = xmlDoc.createElement("FINDADDRESSESPERCUSTOMER")
        
        xmlElement.appendChild xmlCustomerNumber.cloneNode(True)
        xmlElement.appendChild xmlCustomerVersion.cloneNode(True)
        
        Set xmlSubResults = Nothing
        Set xmlSubResults = m_objIDOAssist.FindListMultiple(xmlElement, xmlAddressClassDef)

        If Not xmlSubResults Is Nothing Then
            For Each xmlSubNode In xmlSubResults.selectNodes("FINDADDRESSESPERCUSTOMER")
                xmlChangeNodeName xmlSubNode, "FINDADDRESSESPERCUSTOMER", "ADDRESS"
                xmlSubNode.removeChild xmlSubNode.selectSingleNode("CUSTOMERNUMBER")
                xmlSubNode.removeChild xmlSubNode.selectSingleNode("CUSTOMERVERSIONNUMBER")
                
                xmlCustomer.appendChild xmlSubNode.cloneNode(True)
            Next
        End If
        ' BMIDS00205 End
    Next
    
    Set ICustomerDO_FindCustomersPerApplication = xmlResults
        
    m_objContext.SetComplete
        
FindCustomersPerApplicationVbErr:
    
    'Record not found errors can be ignored here as they are non fatal
    If m_objErrAssist.GetOmigaErrorNumber(Err.Number) = oeRecordNotFound Then
        Resume Next
    End If
    
    
    Set xmlClassDefDoc = Nothing
    Set xmlPhoneClassDef = Nothing
    Set xmlVerificationClassDef = Nothing
    Set xmlAddressClassDef = Nothing    'BMIDS00205
    Set xmlDoc = Nothing
    Set objIClassDef = Nothing
    Set xmlResults = Nothing
    Set xmlSubResults = Nothing
    Set xmlElement = Nothing
    Set xmlCustomerNumber = Nothing
    Set xmlCustomerVersion = Nothing
    Set xmlCustomer = Nothing
    Set xmlSubNode = Nothing
    If Err.Number = 0 Then
        Exit Function
    End If
    
    m_objErrAssist.AddToErrSource (strFunctionName)
    
    If m_objErrAssist.IsSystemError = True Then
        m_objErrAssist.LogError TypeName(Me), strFunctionName, Err.Number, Err.Description
    End If
    
    If m_objErrAssist.IsSystemError Then
        m_objContext.SetAbort
    Else
        m_objContext.SetComplete
    End If
        
    '   re-raise error for business object to interpret as appropriate
    Err.Raise Err.Number, Err.Source, Err.Description
End Function
'BMIDS00026 End

'BMIDS00443
Private Function ICustomerDO_FindLatestCustomerVersion(ByVal vXMLRequest As IXMLDOMElement) As IXMLDOMNode
' header ----------------------------------------------------------------------------------
' description:  Find the lastest version of a customer
' pass:         vxmlRequest     xml element containing the request with CustomerNumber
' return:       xml node containing retrieved data
'------------------------------------------------------------------------------------------
On Error GoTo FindLatestCustomerVersionVbErr

    Const strFunctionName As String = "ICustomerDO_FindLatestCustomerVersion"
    
        
    Dim xmlCustVersClassDefDoc  As FreeThreadedDOMDocument40
    Dim objIClassDef            As IomCustClassDef
    Dim objIADOAssist           As IADOAssist
    Dim xmlCustomerElem         As IXMLDOMElement
    
    Dim adoConnection           As ADODB.Connection
    Dim adoCommand              As ADODB.Command
    Dim adoRecordset            As ADODB.Recordset
    Dim adoParameter            As ADODB.Parameter
    
    Dim strCustomerNumber       As String
    Dim intRetries              As Integer
    Dim intMaxAttempts          As Integer
    Dim intAttempt              As Integer
    Dim blnOpenedOk             As Boolean
    Dim strConnString           As String
    Dim strSQL                  As String
        
    ' Check the correct keys have been passed in
    strCustomerNumber = m_objXmlAssist.GetNodeValue(vXMLRequest, ".//CUSTOMERNUMBER")
    
    Set objIADOAssist = New ADOAssist
    
    Set adoConnection = New ADODB.Connection
    Set adoCommand = New ADODB.Command
    Set adoRecordset = New ADODB.Recordset
    
    strSQL = "SELECT CUSTOMERNUMBER, MAX(CUSTOMERVERSIONNUMBER) CUSTOMERVERSIONNUMBER " & _
             "FROM CUSTOMERVERSION WHERE CUSTOMERNUMBER = ? GROUP BY CUSTOMERNUMBER"

    adoCommand.CommandText = strSQL
    
    ' Set up the parameters to be substituted into the SQL
    Set adoParameter = adoCommand.CreateParameter("CustomerNumber", adBSTR, adParamInput, Len(strCustomerNumber), strCustomerNumber)
    adoCommand.Parameters.Append adoParameter
    
    ' Set Recordset properties. Must be adUseClient for disconnected recordset
    adoRecordset.CursorLocation = adUseClient
    adoRecordset.CursorType = adOpenForwardOnly
    adoRecordset.LockType = adLockReadOnly
    
    intRetries = objIADOAssist.GetRetries()
    intMaxAttempts = 1 + intRetries
    
    blnOpenedOk = False
    intAttempt = 1
    strConnString = objIADOAssist.GetConnStr()
    
    ' Keep trying to open the connection for the specified number of retries
    Do While (Not blnOpenedOk) And (intAttempt <= intMaxAttempts)
    
        adoConnection.Errors.Clear
        adoConnection.Open strConnString
        If adoConnection.Errors.Count = 0 Then
            blnOpenedOk = True
        End If
        intAttempt = intAttempt + 1
    
    Loop
    
    If blnOpenedOk = False Then
        m_objErrAssist.ThrowError TypeName(Me), strFunctionName, oeUnableToConnect
    End If
    
    adoCommand.ActiveConnection = adoConnection
    Set adoRecordset.Source = adoCommand

    'Populate recordset
    adoRecordset.Open
    
    'Disconnect the recordset
    Set adoRecordset.ActiveConnection = Nothing
    adoConnection.Close
    
    If adoRecordset.EOF Then
        m_objErrAssist.ThrowError TypeName(Me), strFunctionName, oeRecordNotFound
    End If
     
    Set objIClassDef = New omCustClassDef
    Set xmlCustVersClassDefDoc = objIClassDef.LoadLatestCustomerVersionData()
       
    ' Convert recordset to XML
    Set xmlCustomerElem = m_objIDOAssist.GetXMLFromRecordSet(adoRecordset, xmlCustVersClassDefDoc)
    Set ICustomerDO_FindLatestCustomerVersion = xmlCustomerElem

    adoRecordset.Close
    
    Set objIClassDef = Nothing
    Set xmlCustVersClassDefDoc = Nothing
    Set objIADOAssist = Nothing
    Set adoConnection = Nothing
    Set adoCommand = Nothing
    Set adoRecordset = Nothing
    Set adoParameter = Nothing
    Set xmlCustomerElem = Nothing
    
    m_objContext.SetComplete
    
    
    Exit Function

FindLatestCustomerVersionVbErr:
    
    
    ' Close the recordset if still open
    If Not adoRecordset Is Nothing Then
        If adoRecordset.State = adStateOpen Then
            adoRecordset.Close
        End If
    End If
    
    ' Close the connection if still open
    If Not adoConnection Is Nothing Then
        If adoConnection.State = adStateOpen Then
            adoConnection.Close
        End If
    End If
    
    Set objIClassDef = Nothing
    Set xmlCustVersClassDefDoc = Nothing
    Set objIADOAssist = Nothing
    Set adoConnection = Nothing
    Set adoCommand = Nothing
    Set adoRecordset = Nothing
    Set adoParameter = Nothing
    Set xmlCustomerElem = Nothing
    
    m_objErrAssist.AddToErrSource (strFunctionName)
    
    If m_objErrAssist.IsSystemError = True Then
        m_objErrAssist.LogError TypeName(Me), strFunctionName, Err.Number, Err.Description
    End If
    
    If m_objErrAssist.IsSystemError Then
        m_objContext.SetAbort
    Else
        m_objContext.SetComplete
    End If
    
    '   re-raise error for business object to interpret as appropriate
    Err.Raise Err.Number, Err.Source, Err.Description

End Function

'BMIDS00443 End

'BMIDS758  Add function

Private Function ICustomerDO_GetRemovedToECustomers(ByVal vXMLRequest As IXMLDOMElement) As IXMLDOMElement
' header ----------------------------------------------------------------------------------
' description:
'   Get the data for the given application number from RemovedTOECustomer table.
' pass:
'   vstrXMLRequest  xml Request data stream containing data which identifies
'                   the criterion to be used in the Query
' return:
'   RemovedToECustomer string containing XML data stream representation
'    of data retrieved in the format
'   <REMOVEDTOECUSTOMERLIST>
'       <REMOVEDTOECUSTOMER>
'           <CUSTOMERNUMBER> ... </CUSTOMERNUMBER>
'       </REMOVEDTOECUSTOMER>
'   </REMOVEDTOECUSTOMERLIST>
' Raise Errors: if record not found, raise omiga4RecordNotFound
'------------------------------------------------------------------------------------------
On Error GoTo ICustomerDO_GetRemovedToECustomersVbErr

    Const strFunctionName As String = "ICustomerDO_GetRemovedToECustomers"

    Dim objErrAssist As New ErrAssist

    Dim objIDOAssist As IDOAssist
    Dim objIClassDef As IomCustClassDef

    Dim objContext As ObjectContext
    Set objContext = GetObjectContext()

    If objContext Is Nothing Then
        Set objIDOAssist = New DOAssist
    Else
        Set objIDOAssist = objContext.CreateInstance(gstrBASE_COMPONENT & ".DOAssist")
    End If

    Set objIClassDef = New omCustClassDef
    Set ICustomerDO_GetRemovedToECustomers = _
            objIDOAssist.FindList(vXMLRequest, objIClassDef.LoadRemovedToECustomer)

    Set objErrAssist = Nothing
    Set objIDOAssist = Nothing
    Set objIClassDef = Nothing
    
    If Not objContext Is Nothing Then
        objContext.SetComplete
        Set objContext = Nothing
    End If

    Exit Function

ICustomerDO_GetRemovedToECustomersVbErr:

    Set objIDOAssist = Nothing
    Set objIClassDef = Nothing

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
        Set objContext = Nothing
    End If

    Set objErrAssist = Nothing

    '   re-raise error for business object to interpret as appropriate
    Err.Raise Err.Number, Err.Source, Err.Description

End Function

'PSC 28/06/2005 MAR5 - Start
Private Function ICustomerDO_GetCustomerAddressListAndRole(ByVal vXMLRequest As IXMLDOMElement _
                    ) As IXMLDOMElement
' header ----------------------------------------------------------------------------------
' description:
'   Get the data from tables CustomerAddress and Address that match the criterion
' pass:
'   vstrXMLRequest  xml Request data stream containing data which identifies
'                   the criterion to be used in the Query
' return:
'   FindCustomerAddressList  string containing XML data stream representation
'    of data retrieved in the format
'   <CUSTOMERADDRESSLIST>
'       <CUSTOMERADDRESS>
'           <CUSTOMERNUMBER> ... </CUSTOMERNUMBER>
'           <CUSTOMERVERSIONNUMBER> .... </CUSTOMERVERSIONNUMBER>
'           ........
'          <CUSTOMERNAME>
'               <CUSTOMERNUMBER> ... </CUSTOMERNUMBER>
'               .......
'               .......
'          </CUSTOMERNAME>
'          <ADDRESS>
'               <ADDRESSGUID> ... </ADDRESSGUID>
'               .......
'         </ADDRESS>
'
'       </CUSTOMERADDRESS>
'  </CUSTOMERADDRESSLIST>
' Raise Errors: if record not found, raise omiga4RecordNotFound
'------------------------------------------------------------------------------------------
On Error GoTo ICustomerDO_GetCustomerAddressListAndRoleVbErr
    Dim strFunctionName As String
    strFunctionName = "ICustomerDO_GetCustomerAddressListAndRole"
    Dim objXmlAssist As New XMLAssist
    Dim objErrAssist As New ErrAssist
    Dim objSQLAssist As New SQLAssist
    Dim objContext As ObjectContext
    Set objContext = GetObjectContext()
    Dim strSQL As String
    Dim strFinalCondition As String ' Condition to be added to select string before fetch data
    Dim strRecCondition As String ' Condition based on single row in xmlRequest
    Dim objIClassDef As IomCustClassDef, objIAddressClassDef As IAddressClassDef
    Dim xmlAddressDef As FreeThreadedDOMDocument40
    Dim xmlCustomerAddressDef As FreeThreadedDOMDocument40
    Dim xmlCustomerNameDef As FreeThreadedDOMDocument40
    Dim xmlCustomerAddressListElem As IXMLDOMElement
    Dim xmlNodeList As IXMLDOMNodeList
    Dim xmlNode As IXMLDOMNode
    Dim objIADOAssist As IADOAssist
    Dim objIDOAssist As IDOAssist
    Dim objIAddressDO As New IAddressDO
    If objContext Is Nothing Then
        Set objIADOAssist = New ADOAssist
        Set objIDOAssist = New DOAssist
    Else
        Set objIADOAssist = objContext.CreateInstance(gstrBASE_COMPONENT & ".ADOAssist")
        Set objIDOAssist = objContext.CreateInstance(gstrBASE_COMPONENT & ".DOAssist")
    End If
    ' Build the criteria to be applied on the view to fetch data
    If vXMLRequest.nodeName = "CUSTOMERADDRESS" Then
        Set xmlNode = vXMLRequest
        strFinalCondition = BuildConditionForFindCustomerAddressList(xmlNode)
    Else
        Set xmlNodeList = vXMLRequest.getElementsByTagName("CUSTOMERADDRESS")
        For Each xmlNode In xmlNodeList
            strRecCondition = BuildConditionForFindCustomerAddressList(xmlNode)
            If strRecCondition <> "" Then
                If strFinalCondition <> "" Then strFinalCondition = strFinalCondition & " OR "
                strFinalCondition = strFinalCondition & "(" & strRecCondition & ")"
            End If
        Next xmlNode
    End If
    
    'Only select Current and Correspondence addresses
    strFinalCondition = "(" & strFinalCondition & ") AND (ADDRESSTYPE = '1' OR ADDRESSTYPE = '2')"
    
    strSQL = "SELECT * FROM FINDCUSTOMERNAMEANDADDRESS"
    If strFinalCondition <> "" Then
        strSQL = strSQL & " Where " & strFinalCondition
    End If
        
    strSQL = strSQL & " Order By ADDRESSTYPE Desc "
    
    'Open the recordset with the SQL string built above
    Dim cmd As New ADODB.Command
    cmd.ActiveConnection = objIADOAssist.GetConnStr
    cmd.CommandText = strSQL
    Dim rstThisRecSet As New ADODB.Recordset
    Set rstThisRecSet = cmd.Execute(, , adCmdText)
    If rstThisRecSet.EOF And rstThisRecSet.BOF Then
        ' raise application error to be interpreted by calling object
        objErrAssist.ThrowError TypeName(Me), strFunctionName, oeRecordNotFound
    End If
    ' Define the classes required for class definitions
    Set objIClassDef = New omCustClassDef
    Set objIAddressClassDef = New AddressClassDef
    Set xmlAddressDef = objIAddressClassDef.LoadAddressData
    Set xmlCustomerAddressDef = objIClassDef.LoadCustomerAddressAndRoleData
    Set xmlCustomerNameDef = objIClassDef.LoadCustomerNameData
    
    Set xmlCustomerAddressListElem = vXMLRequest.ownerDocument.createElement("CUSTOMERADDRESSLIST")
    'Navigate through the record set and add the xml node to list element
    
    Do While Not rstThisRecSet.EOF
        Set xmlNode = objIDOAssist.GetXMLFromRecordSet(rstThisRecSet, xmlCustomerAddressDef, _
                                               xmlCustomerAddressListElem)
        'Add Customer Name data
        Call objIDOAssist.GetXMLFromRecordSet(rstThisRecSet, xmlCustomerNameDef, xmlNode)
        
        'Add Address XML string, If the AddressGuid is not null
        If Not IsNull(rstThisRecSet.Fields("AddressGuid")) Then
            Call objIDOAssist.GetXMLFromRecordSet(rstThisRecSet, xmlAddressDef, xmlNode)
        End If
        rstThisRecSet.MoveNext
    Loop
        rstThisRecSet.Close
    Set ICustomerDO_GetCustomerAddressListAndRole = xmlCustomerAddressListElem
  
    'Clear all the objects
    Set rstThisRecSet = Nothing
    Set cmd = Nothing
    Set objIDOAssist = Nothing
    Set objIADOAssist = Nothing
    Set objIAddressDO = Nothing
    Set objErrAssist = Nothing
    Set objXmlAssist = Nothing
    Set objSQLAssist = Nothing
    Set xmlCustomerAddressListElem = Nothing
    Set xmlNode = Nothing
    Set xmlNodeList = Nothing
    Set objIClassDef = Nothing
    Set objIAddressClassDef = Nothing
    Set xmlAddressDef = Nothing
    Set xmlCustomerAddressDef = Nothing
    Set xmlCustomerNameDef = Nothing
    
    If Not objContext Is Nothing Then
        objContext.SetComplete
        Set objContext = Nothing
    End If
    Exit Function
ICustomerDO_GetCustomerAddressListAndRoleVbErr:
    
    ' Close the recordset if still open
        If Not rstThisRecSet Is Nothing Then
                If rstThisRecSet.State = adStateOpen Then
                        rstThisRecSet.Close
                End If
        End If
    Set rstThisRecSet = Nothing
    Set cmd = Nothing
        Set objIDOAssist = Nothing
    Set objIADOAssist = Nothing
    Set objIAddressDO = Nothing
    Set objXmlAssist = Nothing
    Set objSQLAssist = Nothing
    Set xmlCustomerAddressListElem = Nothing
    Set xmlNode = Nothing
    Set xmlNodeList = Nothing
    Set objIClassDef = Nothing
    Set objIAddressClassDef = Nothing
    Set xmlCustomerNameDef = Nothing
    Set xmlAddressDef = Nothing
    Set xmlCustomerAddressDef = Nothing
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
        Set objContext = Nothing
    End If
    Set objErrAssist = Nothing
    '   re-raise error for business object to interpret as appropriate
    Err.Raise Err.Number, Err.Source, Err.Description
End Function

'PSC 28/06/2005 MAR5 - End
'PSC 10/10/2005 MAR57 - Start
Private Function ICustomerDO_FindCustomerCRSData(ByVal vxmlTableElement As IXMLDOMNode) As IXMLDOMNode
On Error GoTo FindCustomerCRSDataVbErr
    
    Const strFunctionName As String = "ICustomerDO_GetCustomerAddressListAndRole"

    Dim adoConn As ADODB.Connection
    Dim adoCommand As ADODB.Command
    Dim adoStream As ADODB.Stream
    Dim adoParameter As ADODB.Parameter
    Dim objIADOAssist  As IADOAssist
    Dim xmlCustomerData As FreeThreadedDOMDocument40
    Dim xmlCustomersIn As IXMLDOMNodeList
    Dim xmlCustomerIn As IXMLDOMNode
    Dim xmlCustomer As IXMLDOMNode
    Dim xmlChildNodes As IXMLDOMNodeList
    Dim xmlChild As IXMLDOMNode
    Dim xmlResponse As IXMLDOMNode
  
    Dim strData As String
    Dim strTagValue As String

    Set xmlCustomerData = New FreeThreadedDOMDocument40
    xmlCustomerData.validateOnParse = False
    xmlCustomerData.setProperty "NewParser", True
    
    Set xmlResponse = xmlCustomerData.createElement("CUSTOMERS")

    Set adoConn = New ADODB.Connection
    Set adoCommand = New ADODB.Command
    Set adoStream = New ADODB.Stream
    Set objIADOAssist = m_objContext.CreateInstance(gstrBASE_COMPONENT & ".ADOAssist")

    adoConn.ConnectionString = objIADOAssist.GetConnStr()
    
    Set adoParameter = adoCommand.CreateParameter("@CustomerNo", adBSTR, adParamInput)
    adoParameter.Attributes = adParamNullable
    adoCommand.Parameters.Append adoParameter

    Set adoParameter = adoCommand.CreateParameter("@CustomerVersionNo", adInteger, adParamInput)
    adoParameter.Attributes = adParamNullable
    adoCommand.Parameters.Append adoParameter

    Set xmlCustomersIn = vxmlTableElement.selectNodes("CUSTOMERS/CUSTOMER")
    
    For Each xmlCustomerIn In xmlCustomersIn
  
        strTagValue = xmlGetMandatoryNodeText(xmlCustomerIn, "CUSTOMERNUMBER")
        adoCommand.Parameters(0).Size = Len(strTagValue)
        adoCommand.Parameters(0).Value = strTagValue
        
        strTagValue = xmlGetMandatoryNodeText(xmlCustomerIn, "CUSTOMERVERSIONNUMBER")
        adoCommand.Parameters(1).Value = CLng(strTagValue)

        adoConn.Open
        adoStream.Open
                
        With adoCommand
            .CommandText = "USP_GETALLCUSTOMERCRSDATA"
            .CommandType = adCmdStoredProc
            .ActiveConnection = adoConn
            .Properties("Output Stream") = adoStream
            .Properties("XML Root") = "CUSTOMERCRSDATA"
            .Execute , , adExecuteStream
        End With

        strData = adoStream.ReadText()
        
        adoStream.Close
        adoConn.Close
    
        Set xmlCustomerData = m_objXmlAssist.load(strData, TypeName(Me), strFunctionName)
    
        Set xmlCustomer = xmlCustomerData.selectSingleNode("CUSTOMERCRSDATA/CUSTOMER")
        xmlResponse.appendChild xmlCustomer
    
        If Not xmlCustomer Is Nothing Then
            ' Get the phone numbers and move them under the CUSTOMER node
            Set xmlChildNodes = xmlCustomerData.selectNodes("CUSTOMERCRSDATA/PHONENUMBER")
            
            For Each xmlChild In xmlChildNodes
                xmlChild.Attributes.removeNamedItem ("CUSTOMERNUMBER")
                xmlChild.Attributes.removeNamedItem ("CUSTOMERVERSIONNUMBER")
                xmlCustomer.appendChild xmlChild
            Next
        
            ' Get the addresses and move them under the CUSTOMER node
            Set xmlChildNodes = xmlCustomerData.selectNodes("CUSTOMERCRSDATA/ADDRESS")
            
            For Each xmlChild In xmlChildNodes
                xmlChild.Attributes.removeNamedItem ("CUSTOMERNUMBER")
                xmlChild.Attributes.removeNamedItem ("CUSTOMERVERSIONNUMBER")
                xmlCustomer.appendChild xmlChild
            Next
        End If
    Next

    m_objContext.SetComplete
    
    Set ICustomerDO_FindCustomerCRSData = xmlResponse
    
    Set adoConn = Nothing
    Set adoCommand = Nothing
    Set adoStream = Nothing
    Set adoParameter = Nothing
    Set objIADOAssist = Nothing
    Set xmlCustomerData = Nothing
    Set xmlCustomersIn = Nothing
    Set xmlCustomerIn = Nothing
    Set xmlCustomer = Nothing
    Set xmlChildNodes = Nothing
    Set xmlChild = Nothing
    Set xmlResponse = Nothing

    Exit Function

FindCustomerCRSDataVbErr:

    ' Close the stream if still open
    If Not adoStream Is Nothing Then
        If adoStream.State = adStateOpen Then
                adoStream.Close
        End If
    End If
        
    ' Close the connection if still open
    If Not adoConn Is Nothing Then
        If adoConn.State = adStateOpen Then
                adoConn.Close
        End If
    End If
        
    Set adoConn = Nothing
    Set adoCommand = Nothing
    Set adoStream = Nothing
    Set adoParameter = Nothing
    Set objIADOAssist = Nothing
    Set xmlCustomerData = Nothing
    Set xmlCustomersIn = Nothing
    Set xmlCustomerIn = Nothing
    Set xmlCustomer = Nothing
    Set xmlChildNodes = Nothing
    Set xmlChild = Nothing
    Set xmlResponse = Nothing

    m_objErrAssist.AddToErrSource (strFunctionName)
    
    If m_objErrAssist.IsSystemError = True Then
        m_objErrAssist.LogError TypeName(Me), strFunctionName, Err.Number, Err.Description
    End If
    
    m_objContext.SetComplete
  
    '   re-raise error for business object to interpret as appropriate
    Err.Raise Err.Number, Err.Source, Err.Description

End Function
'PSC 10/10/2005 MAR57 - End

'MAR23  Add new function
Private Function ICustomerDO_GetCustomerNumbers(ByVal vXMLRequest As IXMLDOMElement) As IXMLDOMNode
' header ----------------------------------------------------------------------------------
' description:
'   Get the data for all instances of the persistant data associated with
'   this data object for the values supplied
' pass:
'   vxmlTableElement  xml element containing the request
' return:                xml node containing retrieved data
'------------------------------------------------------------------------------------------
On Error GoTo GetCustomerNumbersVbErr

    Const strFunctionName As String = "ICustomerDO_GetCustomerNumbers"
    
    Dim xmlCustomerClassDefDoc As FreeThreadedDOMDocument40
    Dim xmlCustVersClassDefDoc As FreeThreadedDOMDocument40
    Dim xmlOut As FreeThreadedDOMDocument40
    Dim objIClassDef As IomCustClassDef
    Dim objIADOAssist As IADOAssist
    Dim xmlCustomerElem As IXMLDOMElement
    
    Dim adoConnection As ADODB.Connection
    Dim adoCommand As ADODB.Command
    Dim adoRecordset As ADODB.Recordset
    Dim adoParameter As ADODB.Parameter
    
    Dim strOtherSystemCustomerNumber As String
    Dim intRetries As Integer
    Dim intMaxAttempts As Integer
    Dim intAttempt As Integer
    Dim blnOpenedOk As Boolean
    Dim strConnString As String
    
    Dim strCustomerNumber As String
    Dim strApplicationNumber As String
    
    ' Check the correct keys have been passed in
    strApplicationNumber = m_objXmlAssist.GetNodeValue(vXMLRequest, ".//APPLICATIONNUMBER")
    strCustomerNumber = m_objXmlAssist.GetNodeValue(vXMLRequest, ".//CUSTOMERNUMBER")
    
    Set objIADOAssist = New ADOAssist
    
    Set adoConnection = New ADODB.Connection
    Set adoCommand = New ADODB.Command
    Set adoRecordset = New ADODB.Recordset
        
    Dim strSQL As String
    
    strSQL = "SELECT C.CUSTOMERNUMBER, C.OTHERSYSTEMCUSTOMERNUMBER, CR.CUSTOMERVERSIONNUMBER " & _
             "FROM CUSTOMER C, CUSTOMERROLE CR " & _
             "WHERE C.CUSTOMERNUMBER = CR.CUSTOMERNUMBER AND " & _
             "CR.APPLICATIONNUMBER= ? AND " & _
             "CR.CUSTOMERNUMBER= ? "
    
    adoCommand.CommandText = strSQL
    
    ' Set up the parameters to be substituted into the SQL
    Set adoParameter = adoCommand.CreateParameter("ApplicationNumber", adBSTR, adParamInput, Len(strApplicationNumber), strApplicationNumber)
    adoCommand.Parameters.Append adoParameter
    
    Set adoParameter = adoCommand.CreateParameter("CustomerNumber", adBSTR, adParamInput, Len(strCustomerNumber), strCustomerNumber)
    adoCommand.Parameters.Append adoParameter
    
    ' Set Recordset properties. Must be adUseClient for disconnected recordset
    adoRecordset.CursorLocation = adUseClient
    adoRecordset.CursorType = adOpenForwardOnly
    adoRecordset.LockType = adLockReadOnly
    
    intRetries = objIADOAssist.GetRetries()
    intMaxAttempts = 1 + intRetries
    
    blnOpenedOk = False
    intAttempt = 1
    strConnString = objIADOAssist.GetConnStr()
    
    ' Keep trying to open the connection for the specified number of retries
    While (blnOpenedOk = False) And (intAttempt <= intMaxAttempts)
    
        adoConnection.Errors.Clear
        
        adoConnection.Open strConnString
        
        If adoConnection.Errors.Count = 0 Then
            blnOpenedOk = True
        End If
 
        intAttempt = intAttempt + 1
    
    Wend
    
    If blnOpenedOk = False Then
        m_objErrAssist.ThrowError TypeName(Me), strFunctionName, oeUnableToConnect
    End If
    
    adoCommand.ActiveConnection = adoConnection
    Set adoRecordset.Source = adoCommand

    'Populate recordset
    adoRecordset.Open
    
    'Disconnect the recordset
    Set adoRecordset.ActiveConnection = Nothing
    adoConnection.Close
    
    If adoRecordset.EOF Then
        m_objErrAssist.ThrowError TypeName(Me), strFunctionName, oeRecordNotFound
    End If
     
    Set objIClassDef = New omCustClassDef
    Set xmlCustomerClassDefDoc = objIClassDef.LoadCustomerData(omCustVersionNumber)
    Set xmlCustVersClassDefDoc = objIClassDef.LoadCustomerVersionData(omCustVersionNumber)
    
    Set xmlOut = New FreeThreadedDOMDocument40
    xmlOut.validateOnParse = False
    xmlOut.setProperty "NewParser", True
    
    ' Convert recordset to XML
    Set xmlCustomerElem = m_objIDOAssist.GetXMLFromRecordSet(adoRecordset, xmlCustomerClassDefDoc)
    
    m_objIDOAssist.GetXMLFromRecordSet adoRecordset, xmlCustVersClassDefDoc, xmlCustomerElem
    
    Set ICustomerDO_GetCustomerNumbers = xmlCustomerElem

    adoRecordset.Close
    
    
    Set objIClassDef = Nothing
    Set xmlCustomerClassDefDoc = Nothing
    Set xmlCustVersClassDefDoc = Nothing
    Set xmlOut = Nothing
    Set objIADOAssist = Nothing
    Set adoConnection = Nothing
    Set adoCommand = Nothing
    Set adoRecordset = Nothing
    Set adoParameter = Nothing
    Set xmlCustomerElem = Nothing
    
    m_objContext.SetComplete
    
    
    Exit Function

GetCustomerNumbersVbErr:
    
    
    ' Close the recordset if still open
    If Not adoRecordset Is Nothing Then
        If adoRecordset.State = adStateOpen Then
            adoRecordset.Close
        End If
    End If
    
    ' Close the connection if still open
    If Not adoConnection Is Nothing Then
        If adoConnection.State = adStateOpen Then
            adoConnection.Close
        End If
    End If
    
    Set objIClassDef = Nothing
    Set xmlCustomerClassDefDoc = Nothing
    Set xmlCustVersClassDefDoc = Nothing
    Set xmlOut = Nothing
    Set objIADOAssist = Nothing
    Set adoConnection = Nothing
    Set adoCommand = Nothing
    Set adoRecordset = Nothing
    Set adoParameter = Nothing
    Set xmlCustomerElem = Nothing
    
    m_objErrAssist.AddToErrSource (strFunctionName)
    
    If m_objErrAssist.IsSystemError = True Then
        m_objErrAssist.LogError TypeName(Me), strFunctionName, Err.Number, Err.Description
    End If
    
    If m_objErrAssist.IsSystemError Then
        m_objContext.SetAbort
    Else
        m_objContext.SetComplete
    End If
    
    '   re-raise error for business object to interpret as appropriate
    Err.Raise Err.Number, Err.Source, Err.Description

End Function

'MAR23 Add function
Private Sub ICustomerDO_CreateKYCCheck(ByVal vxmlTableElement As IXMLDOMElement)
' header ----------------------------------------------------------------------------------
' description:
'   create an instance of the persistant data associated with this data object
'   for each set of data in the request
' pass:
'   vxmlTableElement  xml element containing the request
' return:                n/a
'------------------------------------------------------------------------------------------
    m_objErrAssist.ThrowError TypeName(Me), "ICustomerDO_CreateKYCCheck", oeNotImplemented

'On Error GoTo CreateKYCCheckVbErr
'
'    Const strFunctionName As String = "ICustomerDO_CreateKYCCheck"
'
'    Dim xmlClassDefDoc As FreeThreadedDOMDocument40
'    Dim objIClassDef As IomCustClassDef
'    Dim xmlDateTime As IXMLDOMElement
'
'    Set xmlDateTime = m_objXmlAssist.GetNode(vxmlTableElement, ".//DATETIME", False)
'
'    If xmlDateTime Is Nothing Then
'        Set xmlDateTime = vxmlTableElement.ownerDocument.createElement("DATETIME")
'        vxmlTableElement.appendChild xmlDateTime
'    End If
'
'    If Len(xmlDateTime.Text) = 0 Then
'        xmlDateTime.Text = Format(Now, "dd/mm/yyyy hh:mm:ss")
'    End If
'
'    Set objIClassDef = New omCustClassDef
'    Set xmlClassDefDoc = objIClassDef.LoadKYCData()
'
'    'Generate a sequence number
'    m_objIDOAssist.GenerateSequenceNumber vxmlTableElement, xmlClassDefDoc, "SEQUENCENUMBER"
'
'    'Create a KnowYourCustomerCheck record
'    m_objIDOAssist.Create vxmlTableElement, xmlClassDefDoc
'
'    Set objIClassDef = Nothing
'    Set xmlClassDefDoc = Nothing
'    Set xmlDateTime = Nothing
'
'    m_objContext.SetComplete
'
'    Exit Sub
'
'CreateKYCCheckVbErr:
'
'    Set objIClassDef = Nothing
'    Set xmlClassDefDoc = Nothing
'
'    m_objErrAssist.AddToErrSource (strFunctionName)
'
'    If m_objErrAssist.IsSystemError = True Then
'        m_objErrAssist.LogError TypeName(Me), strFunctionName, Err.Number, Err.Description
'    End If
'
'    If m_objErrAssist.IsSystemError Then
'        m_objContext.SetAbort
'    Else
'        m_objContext.SetComplete
'    End If
'
'    '   re-raise error for business object to interpret as appropriate
'    Err.Raise Err.Number, Err.Source, Err.Description
    
End Sub


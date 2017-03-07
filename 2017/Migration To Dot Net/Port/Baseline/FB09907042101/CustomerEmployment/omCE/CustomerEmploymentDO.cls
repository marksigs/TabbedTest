VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 3  'UsesTransaction
END
Attribute VB_Name = "CustomerEmploymentDO"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True

'Workfile:      CustomerEmploymentDO.cls
'Copyright:     Copyright © 1999 Marlborough Stirling

'Description:   Code template for CustomerEmployment Data Object
'               contains (implied) standard interface:
'               Public Sub Create(vstrXMLRequest As String)
'               Public Sub Update(vstrXMLRequest As String)
'               Public Sub Delete(vstrXMLRequest As String)
'               Public Function GetData(vstrXMLRequest As String) As String
'               Public Function GetXMLFromRecordSet(vrstRecordSet As ADODB.Recordset) As String
'               Public Function AddDerivedData(vstrXMLData As String) As String
'Dependencies:
'               Add any other dependent components
'
'Issues:        Instancing:         MultiUse
'               MTSTransactionMode: UsesTransaction
'------------------------------------------------------------------------------------------
'History:
'
'Prog   Date     Description
'MCS    09/11/99 Created
'MC     02/06/00 Use INCOMESUMMARYDETAILS view to retrieve Income Summary data
'JLD    13/06/00 bug fix in updateEmployedDetails
' MS    19/06/00 Throw error not raise
'MC     13/07/00 SYS1006 Move SHARESOWNEDINDICATOR to Employment table
'MS     14/07/00    SYS0968 IsSelfEmployed  changed to IsMainSelfEmployed
'SR     24/07/00    New method 'FindEmploymentList'
'MC     31/07/00    New method GetEmploymentAndSelfEmployedDetails
'SR     08/08/00    SYS1385 - Modified method GetEmployedAndSelefEmployment details
'                   Append the UnearnedIncome details, if any
'SR     11/08/00    New method DeleteEarnedIncomeDetails
'                   if appropriate, delete earnedIncome records while deleting of employeddetals
'SR     17/08/00    Modified method 'DeleteEmployment' - Delete the related ThirdParty details
'                   when Employment details are deleted.
'                   Modified private method DeletePreviousEmployments - Do not explicityly call
'                   DeleteThirdParty.
'SR     30/08/00    New method 'GetEmploymentData'
'BG     30/10/00    SYS1600 Changed FindEmploymentAndIncomeDetails to structure XML correctly
'AS     06/03/01    CC012: ICustomerEmploymentDO_DeleteEmploymentDetails now compatible with
'                   Oracle and SQL Server.
'DRC    22/03/01    SYS1932 Added GetAccountantGUID method
'LD     19/06/01    SYS2386 All projects to use guidassist.bas rather than guidassist.cls
'LD     23/06/01    SYS2386 ICustomerEmploymentDO_DeleteEmploymentDetails corrected
'SA     23/07/01    SYS2440 Update statements changed in UpdateContractDetails/UpdateEmployedDetails/
'                   UpdateSelfEmployedDetails/UpdateNetProfitDetails
'AD     05/10/01    Removed MoveFirst as it's a ForwardOnly cursor. SYS 2718.
'MEVA   23/04/02    SYS2742 Add GetTaxGetails
'-----------------------------------------------------------------------------------------------
'BMids History:
'
'Prog   Date        AQR        Description
'MDC    28/06/2002  BMIDS00052 Versions of Employment
'MV     12/08/2002  BMIDS00319 Core AQR SYS1730 Modified ICustomerEmploymentDO_DeleteEmploymentDetails
'SA     19/09/2002  BMIDS00246 New method added - GetEmploymentStatusList'
'SA     17/10/2002  BMIDS00662 GetEmploymentStatusList altered
'SA     12/11/2002  BMIDS00693 DeleteEmploymentDetails - don't raise foreign key constraint error on deleting of third party.
'SA     14/11/2002  BMIDS00929 CreateEmploymentDetails - must create new accountant records too.
'JD     04/10/2004  BMIDS894   DeleteEmploymentDetails - performance enhancements
'GHun   18/10/2004  BMIDS917   Changed GetEmploymentFields to only close recordset if required
'-----------------------------------------------------------------------------------------------
'Epsom History
'
'Prog   Date        AQR        Description
'PE     13/07/2006  EP964      Display an error message if deleting an employment that has related records (Employers or Accountants confirmation).
'-----------------------------------------------------------------------------------------------

Option Explicit

Implements ICustomerEmploymentDO

Public Enum EMPLOYMENTSTATUS
    'These values match the values as defined for the combovalidation table on the database
    Employed 'E
    SelfEmployed ' S
    Contract  'C
    Unemployed 'U
End Enum

Private Sub GetEmploymentFields(ByRef rxmlRequest As IXMLDOMElement, _
                                   ByVal vstrTableName As String)
' header ----------------------------------------------------------------------------------
' description:  Retrieves the value for DATESTARTEDORESTABLISHED from the EMPLOYMENT table
'               for the key values specified in the XML document passed in. That value is
'               then added to the XML document.
'
' pass:         xmlIn         (in/out) DOMDocument
'               vstrTableName (in) String
'                   The name of the table for which the start date is being retrieved for -
'                   i.e. the table from which the data represented by the XML document was retrieved.
'------------------------------------------------------------------------------------------

    On Error GoTo GetEmploymentFieldsVbErr
    
Const strFunctionName As String = "GetEmploymentFields"

Dim rst As ADODB.Recordset, _
    cmd As ADODB.Command

Dim xmlNode     As IXMLDOMNode, _
    xmlRootNode As IXMLDOMNode

Dim objXmlAssist  As XMLAssist, _
    objIAdoAssist As IADOAssist, _
    objErrAssist  As ErrAssist, _
    objSqlAssist  As SQLAssist

Dim strCustomerNumber           As String, _
    strCustomerVersionNumber    As String, _
    strEmploymentSequenceNumber As String
Dim strStartDate      As String, _
    strAccountantGUID As String

Dim strSQL As String
    
    Set objErrAssist = New ErrAssist
    
    '
    ' Initialise
    '
    strStartDate = vbNullString
    Set objIAdoAssist = New ADOAssist

    '
    ' Main
    '
    ' Find table root node
    If rxmlRequest.nodeName = vstrTableName Then
        Set xmlRootNode = rxmlRequest
    Else
        Set xmlRootNode = rxmlRequest.selectSingleNode(".//" & vstrTableName)
    End If

    If xmlRootNode Is Nothing Then
        objErrAssist.ThrowError TypeName(Me), strFunctionName, oeMissingTableName, vstrTableName & " tag not found"
    End If
    
    Set objXmlAssist = New XMLAssist
    strCustomerNumber = objXmlAssist.GetTagValue(xmlRootNode, "CUSTOMERNUMBER")
    strCustomerVersionNumber = objXmlAssist.GetTagValue(xmlRootNode, "CUSTOMERVERSIONNUMBER")
    strEmploymentSequenceNumber = objXmlAssist.GetTagValue(xmlRootNode, "EMPLOYMENTSEQUENCENUMBER")

    Set objSqlAssist = New SQLAssist

    '
    ' Build and execute the SQL query to get the start date
    '
    strSQL = "SELECT DATESTARTEDORESTABLISHED, ACCOUNTANTGUID FROM EMPLOYMENT WHERE " & _
             "CUSTOMERNUMBER = " & objSqlAssist.FormatString(strCustomerNumber) & " and " & _
             "CUSTOMERVERSIONNUMBER = " & strCustomerVersionNumber & " and " & _
             "EMPLOYMENTSEQUENCENUMBER = " & strEmploymentSequenceNumber

    Set cmd = New ADODB.Command
    cmd.ActiveConnection = objIAdoAssist.GetConnStr
    cmd.CommandText = strSQL

    Set rst = cmd.Execute(1, , adCmdText)

    '
    ' Add the value in the recordset to the input XML document
    '
    If Not rst.EOF Then
        
        ' AD 05/10/01 Removed MoveFirst as it's a ForwardOnly cursor. SYS 2718.
        If Not IsNull(rst.Fields.Item(0).Value) Then
            strStartDate = objSqlAssist.DateToString(rst.Fields.Item(0).Value)
        End If
        If Not IsNull(rst.Fields.Item(1).Value) Then
            strAccountantGUID = objSqlAssist.GuidToString(rst.Fields.Item(1).Value)
        End If
    End If

    Set xmlNode = xmlRootNode.appendChild(rxmlRequest.ownerDocument.createElement("DATESTARTEDORESTABLISHED"))
    xmlNode.Text = strStartDate
    Set xmlNode = xmlRootNode.appendChild(rxmlRequest.ownerDocument.createElement("ACCOUNTANTGUID"))
    xmlNode.Text = strAccountantGUID
    rst.Close
    Set rst = Nothing
    
GetEmploymentFieldsExit:
    'rst.Close  'BMIDS917 GHun RecordSet is already closed
    Set objErrAssist = Nothing
    Set xmlNode = Nothing
    Set xmlRootNode = Nothing
    Set objXmlAssist = Nothing
    Set objSqlAssist = Nothing
    Set objIAdoAssist = Nothing
    Set rst = Nothing
    Set cmd = Nothing

    Exit Sub

GetEmploymentFieldsVbErr:
    'BMIDS917 GHun Only close the RecordSet if it is still open
    If Not rst Is Nothing Then
        If rst.State = adStateOpen Then
            rst.Close
        End If
    End If
    'BMIDS917 End
    Set xmlNode = Nothing
    Set xmlRootNode = Nothing
    Set objXmlAssist = Nothing
    Set objSqlAssist = Nothing
    Set objIAdoAssist = Nothing
    Set rst = Nothing
    Set cmd = Nothing

    objErrAssist.AddToErrSource strFunctionName
    
    If objErrAssist.IsSystemError = True Then
        objErrAssist.LogError TypeName(Me), strFunctionName, Err.Number, Err.Description
    End If

    Set objErrAssist = Nothing
    
    Err.Raise Err.Number, Err.Source, Err.Description, Err.HelpFile, Err.HelpContext
End Sub

Private Sub InsertEmploymentSequenceNumber(ByRef rxmlRequest As IXMLDOMElement, _
                                           ByVal vstrEmploymentSequenceNumber As Long)

    On Error GoTo InsertEmploymentSequenceNumberVbErr

Dim objErrAssist As ErrAssist
Const strFunctionName As String = "InsertEmploymentSequenceNumber"

Dim xmlEmploymentSequenceNumber As IXMLDOMNode
     
    Set xmlEmploymentSequenceNumber = rxmlRequest.selectSingleNode("EMPLOYMENTSEQUENCENUMBER")
    If xmlEmploymentSequenceNumber Is Nothing Then
        Set xmlEmploymentSequenceNumber = rxmlRequest.appendChild(rxmlRequest.ownerDocument.createElement("EMPLOYMENTSEQUENCENUMBER"))
    End If
    xmlEmploymentSequenceNumber.Text = vstrEmploymentSequenceNumber

InsertEmploymentSequenceNumberExit:
    On Error Resume Next

    Exit Sub

InsertEmploymentSequenceNumberVbErr:
    On Error Resume Next
    Set objErrAssist = New ErrAssist
    objErrAssist.AddToErrSource strFunctionName
    
    If objErrAssist.IsSystemError = True Then
        objErrAssist.LogError TypeName(Me), strFunctionName, Err.Number, Err.Description
    End If

    Err.Raise Err.Number, Err.Source, Err.Description, Err.HelpFile, Err.HelpContext
End Sub

Private Sub ICustomerEmploymentDO_CreateAccountantDetails(ByVal vxmlTableElement As IXMLDOMElement, _
                                                          ByRef rxmlKeysNode As IXMLDOMNode)
' header ----------------------------------------------------------------------------------
' description:
'   CreateAccountantDetails an instance of the persistant data associated with this data object
'   for each set of data in the request
' pass:
'   vxmlxmlTableElement  xml element containing the request
' return:                n/a
'------------------------------------------------------------------------------------------
On Error GoTo CreateAccountantDetailsVbErr

    Const strFunctionName As String = "ICustomerEmploymentDO_CreateAccountantDetails"
   
    Dim objErrAssist As New ErrAssist
    Dim xmlClassDefDoc As FreeThreadedDOMDocument40
    Dim objIClassDef As IomCEClassDef

    Dim objXmlAssist  As New XMLAssist, _
        objTPAssist As IThirdPartyAssist
    Set objTPAssist = New ThirdPartyAssist
    Dim xmlGUIDNode As IXMLDOMNode
    
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
    ' Set the ACCOUNTANTGUID
    Set xmlGUIDNode = vxmlTableElement.selectSingleNode("ACCOUNTANTGUID")
    If xmlGUIDNode Is Nothing Then
        Set xmlGUIDNode = vxmlTableElement.appendChild(vxmlTableElement.ownerDocument.createElement("ACCOUNTANTGUID"))
    End If
    xmlGUIDNode.Text = CreateGUID()

    Set objIClassDef = New omCEClassDef
    Set xmlClassDefDoc = objIClassDef.LoadAccountantData
    objIDOAssist.Create vxmlTableElement, xmlClassDefDoc
    
    rxmlKeysNode.appendChild xmlGUIDNode

    Set objIClassDef = Nothing
    Set xmlClassDefDoc = Nothing
    Set objIDOAssist = Nothing
    Set objErrAssist = Nothing
    Set objXmlAssist = Nothing
    Set objTPAssist = Nothing
    Set xmlGUIDNode = Nothing

    If Not objContext Is Nothing Then
        objContext.SetComplete
    End If

    Exit Sub

CreateAccountantDetailsVbErr:
    
    Set objIClassDef = Nothing
    Set xmlClassDefDoc = Nothing
    Set objIDOAssist = Nothing
    Set objXmlAssist = Nothing
    Set objTPAssist = Nothing
    Set xmlGUIDNode = Nothing
    
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

Private Sub ICustomerEmploymentDO_CreateContractDetails(ByVal vxmlTableElement As IXMLDOMElement)
' header ----------------------------------------------------------------------------------
' description:
'   CreateContractDetails an instance of the persistant data associated with this data object
'   for each set of data in the request
' pass:
'   vxmlxmlTableElement  xml element containing the request
' return:                n/a
'------------------------------------------------------------------------------------------
On Error GoTo CreateContractDetailsVbErr

    Const strFunctionName As String = "ICustomerEmploymentDO_CreateContractDetails"
   
    Dim objErrAssist As New ErrAssist, _
        objXmlAssist As New XMLAssist
    Dim xmlClassDefDoc As FreeThreadedDOMDocument40
    Dim objIClassDef As IomCEClassDef

    Dim xmlNetProfitXML As IXMLDOMNode
    
    Dim objIDOAssist As IDOAssist
    
    Dim objContext As ObjectContext
    Set objContext = GetObjectContext()
                    
    If objContext Is Nothing Then
        Set objIDOAssist = New DOAssist
    Else
        Set objIDOAssist = objContext.CreateInstance(gstrBASE_COMPONENT & ".DOAssist")
    End If
    
    Set objIClassDef = New omCEClassDef
    Set xmlClassDefDoc = objIClassDef.LoadContractDetailsData()
    objIDOAssist.Create vxmlTableElement, xmlClassDefDoc

    Set xmlNetProfitXML = vxmlTableElement.selectSingleNode(".//NETPROFIT")
    If Not (xmlNetProfitXML Is Nothing) Then
        ' NETPROFIT xml exists, so use it to generate a net profit record
        InsertEmploymentSequenceNumber xmlNetProfitXML, objXmlAssist.GetTagValue(vxmlTableElement, _
            "EMPLOYMENTSEQUENCENUMBER")
        ICustomerEmploymentDO_CreateNetProfitDetails xmlNetProfitXML
    End If
    
    Set objIClassDef = Nothing
    Set xmlClassDefDoc = Nothing
    Set objIDOAssist = Nothing
    Set objErrAssist = Nothing
    Set xmlNetProfitXML = Nothing
    Set objXmlAssist = Nothing

    If Not objContext Is Nothing Then
        objContext.SetComplete
    End If

    Exit Sub

CreateContractDetailsVbErr:
    
    Set objIClassDef = Nothing
    Set xmlClassDefDoc = Nothing
    Set objIDOAssist = Nothing
    Set xmlNetProfitXML = Nothing
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

Private Sub ICustomerEmploymentDO_CreateEarnedIncome(ByVal vxmlTableElement As IXMLDOMElement)
' header ----------------------------------------------------------------------------------
' description:
'   CreateEarnedIncome an instance of the persistant data associated with this data object
'   for each set of data in the request
' pass:
'   vxmlxmlTableElement  xml element containing the request
' return:                n/a
'------------------------------------------------------------------------------------------
On Error GoTo CreateEarnedIncomeVbErr

    Const strFunctionName As String = "ICustomerEmploymentDO_CreateEarnedIncome"
   
    Dim objErrAssist As New ErrAssist
    Dim xmlClassDefDoc As FreeThreadedDOMDocument40
    Dim objIClassDef As IomCEClassDef
    
    Dim objIDOAssist As IDOAssist
    
    Dim objContext As ObjectContext
    Set objContext = GetObjectContext()
                    
    If objContext Is Nothing Then
        Set objIDOAssist = New DOAssist
    Else
        Set objIDOAssist = objContext.CreateInstance(gstrBASE_COMPONENT & ".DOAssist")
    End If
    
    Set objIClassDef = New omCEClassDef
    Set xmlClassDefDoc = objIClassDef.LoadEarnedIncomeData()
    objIDOAssist.GenerateSequenceNumber vxmlTableElement, xmlClassDefDoc, "EARNEDINCOMESEQUENCENUMBER"
    objIDOAssist.Create vxmlTableElement, xmlClassDefDoc
    
    Set objIClassDef = Nothing
    Set xmlClassDefDoc = Nothing
    Set objIDOAssist = Nothing
    Set objErrAssist = Nothing

    If Not objContext Is Nothing Then
        objContext.SetComplete
    End If

    Exit Sub

CreateEarnedIncomeVbErr:
    
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
    
    Set objErrAssist = Nothing
    
    '   re-raise error for business object to interpret as appropriate
    Err.Raise Err.Number, Err.Source, Err.Description
    
End Sub

Private Sub ICustomerEmploymentDO_CreateEmployedDetails(ByVal vxmlTableElement As IXMLDOMElement)
' header ----------------------------------------------------------------------------------
' description:
'   CreateEmployedDetails an instance of the persistant data associated with this data object
'   for each set of data in the request
' pass:
'   vxmlxmlTableElement  xml element containing the request
' return:                n/a
'------------------------------------------------------------------------------------------
On Error GoTo CreateEmployedDetailsVbErr

    Const strFunctionName As String = "ICustomerEmploymentDO_CreateEmployedDetails"
   
    Dim objErrAssist As New ErrAssist, _
        objXmlAssist As New XMLAssist
    Dim xmlClassDefDoc As FreeThreadedDOMDocument40
    Dim objIClassDef As IomCEClassDef

    Dim xmlEarnedIncomeXML As IXMLDOMNode
    Dim xmlNode As IXMLDOMNode
    Dim strEmploymentSequenceNumber As String
    
    Dim objIDOAssist As IDOAssist
    
    Dim objContext As ObjectContext
    Set objContext = GetObjectContext()
                    
    If objContext Is Nothing Then
        Set objIDOAssist = New DOAssist
    Else
        Set objIDOAssist = objContext.CreateInstance(gstrBASE_COMPONENT & ".DOAssist")
    End If
    
    Set objIClassDef = New omCEClassDef
    Set xmlClassDefDoc = objIClassDef.LoadEmployedDetailsData()
    objIDOAssist.Create vxmlTableElement, xmlClassDefDoc

    strEmploymentSequenceNumber = objXmlAssist.GetTagValue(vxmlTableElement, "EMPLOYMENTSEQUENCENUMBER")

    Set xmlEarnedIncomeXML = vxmlTableElement.selectSingleNode(".//EARNEDINCOMELIST")
    If Not (xmlEarnedIncomeXML Is Nothing) Then
        ' NETPROFIT xml exists, so use it to generate a net profit record
        For Each xmlNode In xmlEarnedIncomeXML.childNodes
            InsertEmploymentSequenceNumber xmlNode, strEmploymentSequenceNumber
            ICustomerEmploymentDO_CreateEarnedIncome xmlNode
        Next xmlNode
    End If
    
    'BMIDS00052 MDC 28/06/2002 - Maintain Employment record for SHARESOWNEDINDICATOR
    objXmlAssist.ChangeNodeName vxmlTableElement, "EMPLOYEDDETAILS", "EMPLOYMENT"
    Set xmlClassDefDoc = objIClassDef.LoadEmploymentData(omCEEmployment)
    objIDOAssist.Update vxmlTableElement, xmlClassDefDoc
    'BMIDS00052 MDC 28/06/2002 - End
    
    Set objIClassDef = Nothing
    Set xmlClassDefDoc = Nothing
    Set objIDOAssist = Nothing
    Set objErrAssist = Nothing
    Set xmlEarnedIncomeXML = Nothing
    Set objXmlAssist = Nothing
    Set xmlNode = Nothing

    If Not objContext Is Nothing Then
        objContext.SetComplete
    End If

    Exit Sub

CreateEmployedDetailsVbErr:
    
    Set objIClassDef = Nothing
    Set xmlClassDefDoc = Nothing
    Set objIDOAssist = Nothing
    Set xmlEarnedIncomeXML = Nothing
    Set objXmlAssist = Nothing
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
    End If
    
    Set objErrAssist = Nothing
    
    '   re-raise error for business object to interpret as appropriate
    Err.Raise Err.Number, Err.Source, Err.Description
    
End Sub

Private Sub ICustomerEmploymentDO_CreateEmploymentDetails(ByVal vxmlTableElement As IXMLDOMElement, _
                                                          ByRef rxmlKeysNode As IXMLDOMNode)
' header ----------------------------------------------------------------------------------
' description:
'   CreateEmploymentDetails an instance of the persistant data associated with this data object
'   for each set of data in the request
' pass:
'   vxmlxmlTableElement  xml element containing the request
' return:                n/a
'------------------------------------------------------------------------------------------
On Error GoTo CreateEmploymentDetailsVbErr

    Const strFunctionName As String = "ICustomerEmploymentDO_CreateEmploymentDetails"
   
    Dim objErrAssist As New ErrAssist
    Dim xmlClassDefDoc As FreeThreadedDOMDocument40
    Dim objIClassDef As IomCEClassDef
    Dim lngEmploymentSequenceNumber As Long
    Dim strEmploymentStatus As String
    Dim xmlDetailsXML As IXMLDOMNode
    Dim objComboDO As New ComboDO
    Dim xmlAccountantNode As IXMLDOMNode        'BMIDS00929 SA 14/11/02
    Dim strAccountantGUID As String             'BMIDS00929 SA 14/11/02

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
    
    'BMIDS00929 SA 14/11/2002 Need to maintain the Accountant table too. {
    If Trim$(objXmlAssist.GetTagValue(vxmlTableElement, "ACCOUNTANT")) <> "" Then
        'Check if we have an accountant - if so, create a new accountant record
        Set xmlAccountantNode = vxmlTableElement.selectSingleNode(".//ACCOUNTANT")
        If Not xmlAccountantNode Is Nothing Then
            'Create accountant - this in itself maintains third parties
            ICustomerEmploymentDO_CreateAccountantDetails xmlAccountantNode, rxmlKeysNode
            strAccountantGUID = Trim$(objXmlAssist.GetTagValue(rxmlKeysNode, "ACCOUNTANTGUID"))
        End If
    End If
    'BMIDS00929 }
    
    '
    ' Maintain EMPLOYMENT table
    '
    Set objIClassDef = New omCEClassDef
    Set xmlClassDefDoc = objIClassDef.LoadEmploymentData
    lngEmploymentSequenceNumber = objIDOAssist.GenerateSequenceNumber(vxmlTableElement, xmlClassDefDoc, _
        "EMPLOYMENTSEQUENCENUMBER")
        
    'BMIDS00929 SA 14/11/2002 Need to update the AccountantGuid on the employment table if we have created a new accountant record.
    If strAccountantGUID <> "" Then
        Set xmlAccountantNode = vxmlTableElement.selectSingleNode("ACCOUNTANTGUID")
        If Not xmlAccountantNode Is Nothing Then
            xmlAccountantNode.Text = strAccountantGUID
        End If
    End If
    'BMIDS00929 }
    objIDOAssist.Create vxmlTableElement, xmlClassDefDoc

    '
    ' Create the appropriate EMPLOYEDDETAILS, SELFEMPLOYEDDETAILS or CONTRACTDETAILS record
    '
    strEmploymentStatus = Trim$(objXmlAssist.GetTagValue(vxmlTableElement, "EMPLOYMENTSTATUS"))
    If strEmploymentStatus <> "" Then
        strEmploymentStatus = objComboDO.GetFirstComboValidation("EmploymentStatus", strEmploymentStatus)
    End If

    Select Case strEmploymentStatus
        Case "E" ' EMPLOYEDDETAILS
            Set xmlDetailsXML = vxmlTableElement.selectSingleNode("EMPLOYEDDETAILS")
            If Not (xmlDetailsXML Is Nothing) Then
                InsertEmploymentSequenceNumber xmlDetailsXML, CStr(lngEmploymentSequenceNumber)
                ICustomerEmploymentDO_CreateEmployedDetails xmlDetailsXML
            End If
        Case "S" ' SELFEMPLOYEDDETAILS
            Set xmlDetailsXML = vxmlTableElement.selectSingleNode("SELFEMPLOYEDDETAILS")
            If Not (xmlDetailsXML Is Nothing) Then
                InsertEmploymentSequenceNumber xmlDetailsXML, CStr(lngEmploymentSequenceNumber)
                ICustomerEmploymentDO_CreateSelfEmployedDetails xmlDetailsXML
            End If
        Case "C" ' CONTRACTDETAILS
            Set xmlDetailsXML = vxmlTableElement.selectSingleNode("CONTRACTDETAILS")
            If Not (xmlDetailsXML Is Nothing) Then
                InsertEmploymentSequenceNumber xmlDetailsXML, CStr(lngEmploymentSequenceNumber)
                ICustomerEmploymentDO_CreateContractDetails xmlDetailsXML
            End If
        Case Else
            ' No need to generate any extra records for employment statuses of this type
    End Select

    rxmlKeysNode.appendChild(rxmlKeysNode.ownerDocument.createElement("EMPLOYMENTSEQUENCENUMBER")).Text = _
        lngEmploymentSequenceNumber
    
    Set objIClassDef = Nothing
    Set xmlClassDefDoc = Nothing
    Set objIDOAssist = Nothing
    Set objErrAssist = Nothing
    Set objXmlAssist = Nothing
    Set objTPAssist = Nothing
    Set xmlDetailsXML = Nothing
    Set objComboDO = Nothing
    Set xmlAccountantNode = Nothing 'BMIDS00929 SA 14/11/02
    
    If Not objContext Is Nothing Then
        objContext.SetComplete
    End If

    Exit Sub

CreateEmploymentDetailsVbErr:
    
    Set objIClassDef = Nothing
    Set xmlClassDefDoc = Nothing
    Set objIDOAssist = Nothing
    Set objXmlAssist = Nothing
    Set objTPAssist = Nothing
    Set xmlDetailsXML = Nothing
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
    End If
    
    Set objErrAssist = Nothing
    
    '   re-raise error for business object to interpret as appropriate
    Err.Raise Err.Number, Err.Source, Err.Description
    
End Sub

Private Sub ICustomerEmploymentDO_CreateNetProfitDetails(ByVal vxmlTableElement As IXMLDOMElement)
' header ----------------------------------------------------------------------------------
' description:
'   CreateNetProfitDetails an instance of the persistant data associated with this data object
'   for each set of data in the request
' pass:
'   vxmlxmlTableElement  xml element containing the request
' return:                n/a
'------------------------------------------------------------------------------------------
On Error GoTo CreateNetProfitDetailsVbErr

    Const strFunctionName As String = "ICustomerEmploymentDO_CreateNetProfitDetails"
   
    Dim objErrAssist As New ErrAssist
    Dim xmlClassDefDoc As FreeThreadedDOMDocument40
    Dim objIClassDef As IomCEClassDef
    
    Dim objIDOAssist As IDOAssist
    
    Dim objContext As ObjectContext
    Set objContext = GetObjectContext()
                    
    If objContext Is Nothing Then
        Set objIDOAssist = New DOAssist
    Else
        Set objIDOAssist = objContext.CreateInstance(gstrBASE_COMPONENT & ".DOAssist")
    End If
    
    Set objIClassDef = New omCEClassDef
    Set xmlClassDefDoc = objIClassDef.LoadNetProfitData()
    objIDOAssist.Create vxmlTableElement, xmlClassDefDoc
    
    Set objIClassDef = Nothing
    Set xmlClassDefDoc = Nothing
    Set objIDOAssist = Nothing
    Set objErrAssist = Nothing

    If Not objContext Is Nothing Then
        objContext.SetComplete
    End If

    Exit Sub

CreateNetProfitDetailsVbErr:
    
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
    
    Set objErrAssist = Nothing
    
    '   re-raise error for business object to interpret as appropriate
    Err.Raise Err.Number, Err.Source, Err.Description
    
End Sub

Private Sub ICustomerEmploymentDO_CreateOtherIncome(ByVal vxmlTableElement As IXMLDOMElement)
' header ----------------------------------------------------------------------------------
' description:
'   CreateOtherIncome an instance of the persistant data associated with this data object
'   for each set of data in the request
' pass:
'   vxmlxmlTableElement  xml element containing the request
' return:                n/a
'------------------------------------------------------------------------------------------
On Error GoTo CreateOtherIncomeVbErr

    Const strFunctionName As String = "ICustomerEmploymentDO_CreateOtherIncome"
   
    Dim objErrAssist As New ErrAssist
    Dim xmlClassDefDoc As FreeThreadedDOMDocument40
    Dim objIClassDef As IomCEClassDef
    
    Dim objIDOAssist As IDOAssist
    
    Dim objContext As ObjectContext
    Set objContext = GetObjectContext()
                    
    If objContext Is Nothing Then
        Set objIDOAssist = New DOAssist
    Else
        Set objIDOAssist = objContext.CreateInstance(gstrBASE_COMPONENT & ".DOAssist")
    End If
    
    Set objIClassDef = New omCEClassDef
    Set xmlClassDefDoc = objIClassDef.LoadUnearnedIncomeData()
    objIDOAssist.GenerateSequenceNumber vxmlTableElement, xmlClassDefDoc, "UNEARNEDINCOMESEQUENCENUMBER"
    objIDOAssist.Create vxmlTableElement, xmlClassDefDoc
    
    Set objIClassDef = Nothing
    Set xmlClassDefDoc = Nothing
    Set objIDOAssist = Nothing
    Set objErrAssist = Nothing

    If Not objContext Is Nothing Then
        objContext.SetComplete
    End If

    Exit Sub

CreateOtherIncomeVbErr:
    
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
    
    Set objErrAssist = Nothing
    
    '   re-raise error for business object to interpret as appropriate
    Err.Raise Err.Number, Err.Source, Err.Description
    
End Sub

Private Sub ICustomerEmploymentDO_CreateSelfEmployedDetails(ByVal vxmlTableElement As IXMLDOMElement)
' header ----------------------------------------------------------------------------------
' description:
'   CreateSelfEmployedDetails an instance of the persistant data associated with this data object
'   for each set of data in the request
' pass:
'   vxmlxmlTableElement  xml element containing the request
' return:                n/a
'------------------------------------------------------------------------------------------
On Error GoTo CreateSelfEmployedDetailsVbErr

    Const strFunctionName As String = "ICustomerEmploymentDO_CreateSelfEmployedDetails"
   
    Dim objErrAssist As New ErrAssist, _
        objXmlAssist As New XMLAssist
    Dim xmlClassDefDoc As FreeThreadedDOMDocument40
    Dim objIClassDef As IomCEClassDef

    Dim xmlNetProfitXML As IXMLDOMNode
    
    Dim objIDOAssist As IDOAssist
    
    Dim objContext As ObjectContext
    Set objContext = GetObjectContext()
                    
    If objContext Is Nothing Then
        Set objIDOAssist = New DOAssist
    Else
        Set objIDOAssist = objContext.CreateInstance(gstrBASE_COMPONENT & ".DOAssist")
    End If
    
    Set objIClassDef = New omCEClassDef
    Set xmlClassDefDoc = objIClassDef.LoadSelfEmployedDetailsData()
    objIDOAssist.Create vxmlTableElement, xmlClassDefDoc

    Set xmlNetProfitXML = vxmlTableElement.selectSingleNode(".//NETPROFIT")
    If Not (xmlNetProfitXML Is Nothing) Then
        ' NETPROFIT xml exists, so use it to generate a net profit record
        InsertEmploymentSequenceNumber xmlNetProfitXML, objXmlAssist.GetTagValue(vxmlTableElement, _
            "EMPLOYMENTSEQUENCENUMBER")
        ICustomerEmploymentDO_CreateNetProfitDetails xmlNetProfitXML
    End If
    
    Set objIClassDef = Nothing
    Set xmlClassDefDoc = Nothing
    Set objIDOAssist = Nothing
    Set objErrAssist = Nothing
    Set xmlNetProfitXML = Nothing
    Set objXmlAssist = Nothing

    If Not objContext Is Nothing Then
        objContext.SetComplete
    End If

    Exit Sub

CreateSelfEmployedDetailsVbErr:
    
    Set objIClassDef = Nothing
    Set xmlClassDefDoc = Nothing
    Set objIDOAssist = Nothing
    Set xmlNetProfitXML = Nothing
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

Private Sub ICustomerEmploymentDO_CreateTaxDetails(ByVal vxmlTableElement As IXMLDOMElement)
' header ----------------------------------------------------------------------------------
' description:
'   CreateTaxDetails an instance of the persistant data associated with this data object
'   for each set of data in the request
' pass:
'   vxmlxmlTableElement  xml element containing the request
' return:                n/a
'------------------------------------------------------------------------------------------
On Error GoTo CreateTaxDetailsVbErr

    Const strFunctionName As String = "ICustomerEmploymentDO_CreateTaxDetails"
   
    Dim objErrAssist As New ErrAssist
    Dim xmlClassDefDoc As FreeThreadedDOMDocument40
    Dim objIClassDef As IomCEClassDef
    
    Dim objIDOAssist As IDOAssist
    
    Dim objContext As ObjectContext
    Set objContext = GetObjectContext()
                    
    If objContext Is Nothing Then
        Set objIDOAssist = New DOAssist
    Else
        Set objIDOAssist = objContext.CreateInstance(gstrBASE_COMPONENT & ".DOAssist")
    End If
    
    Set objIClassDef = New omCEClassDef
    Set xmlClassDefDoc = objIClassDef.LoadIncomeSummaryData()
    objIDOAssist.Create vxmlTableElement, xmlClassDefDoc
    
    Set objIClassDef = Nothing
    Set xmlClassDefDoc = Nothing
    Set objIDOAssist = Nothing
    Set objErrAssist = Nothing

    If Not objContext Is Nothing Then
        objContext.SetComplete
    End If

    Exit Sub

CreateTaxDetailsVbErr:
    
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
    
    Set objErrAssist = Nothing
    
    '   re-raise error for business object to interpret as appropriate
    Err.Raise Err.Number, Err.Source, Err.Description
    
End Sub

Private Sub ICustomerEmploymentDO_DeleteAccountantDetails(ByVal vxmlTableElement As IXMLDOMElement)
' header ----------------------------------------------------------------------------------
' description:
'   DeleteAccountantDetails a single instance of the persistant data associated with this
'   data object
' pass:
'   vxmlxmlTableElement  xml element containing the request
' return:                n/a
'------------------------------------------------------------------------------------------
On Error GoTo DeleteAccountantDetailsVbErr

    Const strFunctionName As String = "ICustomerEmploymentDO_DeleteAccountantDetails"

    Dim objErrAssist As New ErrAssist
    
    Dim objContext As ObjectContext
    Set objContext = GetObjectContext()
            
    Dim objIDOAssist As IDOAssist, _
        objXmlAssist As New XMLAssist, _
        objTPAssist As IThirdPartyAssist
    Set objTPAssist = New ThirdPartyAssist

    Dim xmlClassDefDoc As FreeThreadedDOMDocument40
    Dim objIClassDef As IomCEClassDef

    Dim xmlData As IXMLDOMNode

    Dim strThirdPartyGUID As String
    
    If objContext Is Nothing Then
        Set objIDOAssist = New DOAssist
    Else
        Set objIDOAssist = objContext.CreateInstance(gstrBASE_COMPONENT & ".DOAssist")
    End If
            
    Set objIClassDef = New omCEClassDef
    Set xmlClassDefDoc = objIClassDef.LoadAccountantData

    ' Retrieve THIRDPARTYGUID
    Set xmlData = ICustomerEmploymentDO_GetAccountantDetails(vxmlTableElement)
    strThirdPartyGUID = objXmlAssist.GetTagValue(xmlData, "THIRDPARTYGUID")

    objIDOAssist.Delete vxmlTableElement, xmlClassDefDoc
    
    If strThirdPartyGUID <> "" Then
        ' Delete the thirdparty details
        objTPAssist.DeleteThirdParty strThirdPartyGUID
    End If

    Set objIClassDef = Nothing
    Set xmlClassDefDoc = Nothing
    Set objIDOAssist = Nothing
    Set objErrAssist = Nothing
    Set objXmlAssist = Nothing
    Set objTPAssist = Nothing
    
    If Not objContext Is Nothing Then
        objContext.SetComplete
    End If
    
    Exit Sub

DeleteAccountantDetailsVbErr:

    Set objIClassDef = Nothing
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

Private Sub ICustomerEmploymentDO_DeleteEarnedIncome(ByVal vxmlTableElement As IXMLDOMElement)
' header ----------------------------------------------------------------------------------
' description:
' pass:
'   vxmlxmlTableElement  xml element containing the request
' return:                n/a
'------------------------------------------------------------------------------------------
On Error GoTo DeleteEarnedIncomeVbErr

    Const strFunctionName As String = "ICustomerEmploymentDO_DeleteEarnedIncome"
   
    Dim objErrAssist As New ErrAssist
    Dim xmlClassDefDoc As FreeThreadedDOMDocument40
    Dim objIClassDef As IomCEClassDef
    
    Dim objIDOAssist As IDOAssist
    
    Dim objContext As ObjectContext
    Set objContext = GetObjectContext()
                    
    If objContext Is Nothing Then
        Set objIDOAssist = New DOAssist
    Else
        Set objIDOAssist = objContext.CreateInstance(gstrBASE_COMPONENT & ".DOAssist")
    End If
    
    Set objIClassDef = New omCEClassDef
    Set xmlClassDefDoc = objIClassDef.LoadEarnedIncomeData()
    objIDOAssist.Delete vxmlTableElement, xmlClassDefDoc
    
    Set objIClassDef = Nothing
    Set xmlClassDefDoc = Nothing
    Set objIDOAssist = Nothing
    Set objErrAssist = Nothing

    If Not objContext Is Nothing Then
        objContext.SetComplete
    End If

    Exit Sub

DeleteEarnedIncomeVbErr:
    
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
    
    Set objErrAssist = Nothing
    
    '   re-raise error for business object to interpret as appropriate
    Err.Raise Err.Number, Err.Source, Err.Description
End Sub

Private Sub ICustomerEmploymentDO_DeleteEmploymentDetails(ByVal vxmlTableElement As IXMLDOMElement)
    
    On Error GoTo DeleteEmploymentDetailsVbErr

    Const strFunctionName As String = "ICustomerEmploymentDO_DeleteEmploymentDetails"

    Dim objErrAssist As New ErrAssist
    
    Dim objContext As ObjectContext
    Set objContext = GetObjectContext()
            
    Dim objIAdoAssist As IADOAssist, _
        objSqlAssist As New SQLAssist, _
        objXmlAssist As New XMLAssist, _
        objIThirdPartyAssist As IThirdPartyAssist

    Dim strCustomerNumber           As String, _
        strCustomerVersionNumber    As String, _
        strEmploymentSequenceNumber As String, _
        strThirdPartyGUID As String
        
    Dim vntAccountantGUID As Variant
    Dim vntThirdPartyGUID As Variant
    Dim xmlAccountantNode As IXMLDOMNode
    Dim xmlEmployment As IXMLDOMNode
    Dim cmd As ADODB.Command

    If objContext Is Nothing Then
        Set objIAdoAssist = New ADOAssist
        Set objIThirdPartyAssist = New ThirdPartyAssist
    Else
        Set objIAdoAssist = objContext.CreateInstance(gstrBASE_COMPONENT & ".ADOAssist")
        Set objIThirdPartyAssist = objContext.CreateInstance(App.Title & ".ThirdPartyAssist")
    End If
            
    strCustomerNumber = objXmlAssist.GetTagValue(vxmlTableElement, "CUSTOMERNUMBER")
    strCustomerVersionNumber = objXmlAssist.GetTagValue(vxmlTableElement, "CUSTOMERVERSIONNUMBER")
    strEmploymentSequenceNumber = objXmlAssist.GetTagValue(vxmlTableElement, "EMPLOYMENTSEQUENCENUMBER")
    
    ' Get the ThirdPartyGuid, it is required later.
    ' if it mentioned in the request, get it else from the database
    strThirdPartyGUID = objXmlAssist.GetTagValue(vxmlTableElement, "THIRDPARTYGUID")
    'Set objIClassDef = New omCEClassDef
    'Set xmlClassDef = objIClassDef.LoadEmploymentData()
    
    'JD  BMIDS894 ThirdPartyGUID now returned from the StoredProcedure
    'If strThirdPartyGUID = "" Then
    '    Set xmlEmployment = objIDOAssist.GetData(vxmlTableElement, xmlClassDef)
    '    strThirdPartyGUID = xmlEmployment.selectSingleNode("THIRDPARTYGUID").Text
    'End If

    If (strCustomerNumber = "") Or _
       (strCustomerVersionNumber = "") Or _
       (strEmploymentSequenceNumber = "") Then
        objErrAssist.ThrowError TypeName(Me), strFunctionName, oeMissingPrimaryTag
    End If
    
    'EP964 - Peter Edney - 13/07/2006
    Set cmd = New ADODB.Command
    With cmd
        .ActiveConnection = objIAdoAssist.GetConnStr
        .CommandType = adCmdStoredProc
        .CommandText = "usp_checkemployment"
        .Parameters.Append .CreateParameter("strCUSTOMERNUMBER", adVarChar, adParamInput, Len(strCustomerNumber), strCustomerNumber)
        .Parameters.Append .CreateParameter("intCUSTOMERVERSIONNUMBER", adInteger, adParamInput, , CInt(strCustomerVersionNumber))
        .Parameters.Append .CreateParameter("intEMPLOYMENTSEQUENCENUMBER", adInteger, adParamInput, , CInt(strEmploymentSequenceNumber))
        .Parameters.Append .CreateParameter("strFOREIGNTABLE", adVarChar, adParamOutput, 2556)
        .Execute , , adExecuteNoRecords
        
        If Not IsNull(.Parameters("strFOREIGNTABLE").Value) Then
            Select Case LCase(.Parameters("strFOREIGNTABLE").Value)
            Case "employersreference"
                Err.Raise 1, , "Employment details cannot be deleted as an employers confirmation has been created." & vbCrLf & "---"
            Case "accountantsreference"
                Err.Raise 1, , "Employment details cannot be deleted as an accountants confirmation has been created." & vbCrLf & "---"
            End Select
        End If
    
    End With
    
    Set cmd = New ADODB.Command
    With cmd
        .ActiveConnection = objIAdoAssist.GetConnStr
        .CommandType = adCmdStoredProc
       'SYS1730 - Use single stored proc calling convention. Do not use Parameters.Refresh method.
'    #If GENERIC_SQL Then
        ' AS 06/03/01 CC012: Compatible with Oracle and SQL Server.
        ' DELETEEMPLOYMENT stored procedure replaced by DELETEEMPLOYMENT_GS.
        .CommandText = "DELETEEMPLOYMENT_GS"
        .Parameters.Append .CreateParameter("strCUSTOMERNUMBER", adVarChar, adParamInput, Len(strCustomerNumber), strCustomerNumber)
        .Parameters.Append .CreateParameter("intCUSTOMERVERSIONNUMBER", adInteger, adParamInput, , CInt(strCustomerVersionNumber))
        .Parameters.Append .CreateParameter("intEMPLOYMENTSEQUENCENUMBER", adInteger, adParamInput, , CInt(strEmploymentSequenceNumber))
        .Parameters.Append .CreateParameter("guidACCOUNTANTGUID", adBinary, adParamOutput, 16)
        .Parameters.Append .CreateParameter("guidTHIRDPARTYGUID", adBinary, adParamOutput, 16)  'JD BMIDS894
        
        'EP964 - Peter Edney - 13/07/2006
        '.Execute , , adCmdStoredProc
        .Execute , , adExecuteNoRecords
        
        If Not IsNull(.Parameters("guidACCOUNTANTGUID").Value) Then
            vntAccountantGUID = objSqlAssist.GuidToString(.Parameters("guidACCOUNTANTGUID").Value)
        End If
        'JD BMIDS894
        If Not IsNull(.Parameters("guidTHIRDPARTYGUID").Value) Then
            vntThirdPartyGUID = objSqlAssist.GuidToString(.Parameters("guidTHIRDPARTYGUID").Value)
            If strThirdPartyGUID = "" Then
                strThirdPartyGUID = CStr(vntThirdPartyGUID)
            End If
        End If
        
'    #Else
'        .CommandText = "DELETEEMPLOYMENT"
'        .Parameters.Refresh
'        .Parameters("strCUSTOMERNUMBER").Value = strCustomerNumber
'        .Parameters("intCUSTOMERVERSIONNUMBER").Value = strCustomerVersionNumber
'        .Parameters("intEMPLOYMENTSEQUENCENUMBER").Value = strEmploymentSequenceNumber
'        .Execute , , adCmdStoredProc
'        vntAccountantGUID = Trim(.Parameters("strACCOUNTANTGUID").Value)
'    #End If
    'SYS1730 - End.
    End With
    
    If Not IsNull(vntAccountantGUID) Then
        If CStr(vntAccountantGUID) <> "" Then
            ' Delete the corresponding accountant record
            Set xmlAccountantNode = vxmlTableElement.ownerDocument.createElement("ACCOUNTANT")
            xmlAccountantNode.appendChild(xmlAccountantNode.ownerDocument.createElement("ACCOUNTANTGUID")).Text = _
                CStr(vntAccountantGUID)
    
            ICustomerEmploymentDO_DeleteAccountantDetails xmlAccountantNode
        End If
    End If
    
    ' SR : SYS1257 - Delete the related ThirdParty, if any
    ' SA BMIDS00693 Because of the intrduction of customer versioning, the employment
    ' table now holds more than one record with the same third party guid, so a foreign key constraint
    ' error was occurring. Need to prevent this as it's okay.
    If strThirdPartyGUID <> "" Then
        On Error GoTo ThirdPartyForeignKeyConstraintErr        '++ BMIDS00693
        objIThirdPartyAssist.DeleteThirdParty strThirdPartyGUID
        On Error GoTo DeleteEmploymentDetailsVbErr  '++ BMIDS00693 reset error handling
    End If
    
    Set objIThirdPartyAssist = Nothing
    Set objIAdoAssist = Nothing
    Set objErrAssist = Nothing
    Set cmd = Nothing
    Set objSqlAssist = Nothing
    Set objXmlAssist = Nothing
    Set xmlAccountantNode = Nothing
    
    If Not objContext Is Nothing Then
        objContext.SetComplete
    End If
    
    Exit Sub

'BMIDS00693 Okay to let foreign key constraint errors on third party table go through {
ThirdPartyForeignKeyConstraintErr:
    If Err.Number = -2147217873 Then
        Resume Next
    End If
'BMIDS00693 }
DeleteEmploymentDetailsVbErr:


    Set objIAdoAssist = Nothing
    Set cmd = Nothing
    Set objSqlAssist = Nothing
    Set objXmlAssist = Nothing
    Set xmlAccountantNode = Nothing
    
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

Private Sub ICustomerEmploymentDO_DeleteOtherIncome(ByVal vxmlTableElement As IXMLDOMElement)
' header ----------------------------------------------------------------------------------
' description:
'   DeleteOtherIncome a single instance of the persistant data associated with this
'   data object
' pass:
'   vxmlxmlTableElement  xml element containing the request
' return:                n/a
'------------------------------------------------------------------------------------------
On Error GoTo DeleteOtherIncomeVbErr

    Const strFunctionName As String = "ICustomerEmploymentDO_DeleteOtherIncome"

    Dim objErrAssist As New ErrAssist
    
    Dim objContext As ObjectContext
    Set objContext = GetObjectContext()
            
    Dim objIDOAssist As IDOAssist
    Dim xmlClassDefDoc As FreeThreadedDOMDocument40
    Dim objIClassDef As IomCEClassDef
    
    If objContext Is Nothing Then
        Set objIDOAssist = New DOAssist
    Else
        Set objIDOAssist = objContext.CreateInstance(gstrBASE_COMPONENT & ".DOAssist")
    End If
            
    Set objIClassDef = New omCEClassDef
    Set xmlClassDefDoc = objIClassDef.LoadUnearnedIncomeData()
    objIDOAssist.Delete vxmlTableElement, xmlClassDefDoc
    
    Set objIClassDef = Nothing
    Set xmlClassDefDoc = Nothing
    Set objIDOAssist = Nothing
    Set objErrAssist = Nothing
    
    If Not objContext Is Nothing Then
        objContext.SetComplete
    End If
    
    Exit Sub

DeleteOtherIncomeVbErr:

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
    
    Set objErrAssist = Nothing
    '   re-raise error for business object to interpret as appropriate
    Err.Raise Err.Number, Err.Source, Err.Description
End Sub

Private Function ICustomerEmploymentDO_FindEarnedIncomeList(ByVal vxmlTableElement As IXMLDOMElement) As IXMLDOMNode
' header ----------------------------------------------------------------------------------
' description:
'   Get the data for all instances of the persistant data associated with
'   this data object for the values supplied
' pass:
'   vxmlxmlTableElement  xml element containing the request
' return:                xml node containing retrieved data
'------------------------------------------------------------------------------------------
On Error GoTo FindEarnedIncomeListVbErr

    Const strFunctionName As String = "ICustomerEmploymentDO_FindEarnedIncomeList"

    Dim objErrAssist As New ErrAssist
    
    Dim objIDOAssist As IDOAssist, _
        objXmlAssist As New XMLAssist
    Dim xmlClassDefDoc As FreeThreadedDOMDocument40
    Dim objIClassDef As IomCEClassDef

    Dim strCustomerNumber           As String, _
        strCustomerVersionNumber    As String, _
        strEmploymentSequenceNumber As String
    Dim xmlEmploymentRequest    As IXMLDOMNode, _
        xmlEmploymentResponse   As IXMLDOMNode, _
        xmlNetMonthlyIncomeNode As IXMLDOMNode
    Dim xmlReturn As IXMLDOMNode
    
    Dim objContext As ObjectContext
    Set objContext = GetObjectContext()
         
    If objContext Is Nothing Then
        Set objIDOAssist = New DOAssist
    Else
        Set objIDOAssist = objContext.CreateInstance(gstrBASE_COMPONENT & ".DOAssist")
    End If
   
    strCustomerNumber = objXmlAssist.GetTagValue(vxmlTableElement, "CUSTOMERNUMBER")
    strCustomerVersionNumber = objXmlAssist.GetTagValue(vxmlTableElement, "CUSTOMERVERSIONNUMBER")
    strEmploymentSequenceNumber = objXmlAssist.GetTagValue(vxmlTableElement, "EMPLOYMENTSEQUENCENUMBER")
     
    Set objIClassDef = New omCEClassDef
    Set xmlClassDefDoc = objIClassDef.LoadEarnedIncomeData
    Set xmlReturn = objIDOAssist.FindList(vxmlTableElement, xmlClassDefDoc)

    ' Retrieve NETMONTHLYINCOME from the employment table
    Set xmlEmploymentRequest = vxmlTableElement.ownerDocument.createElement("EMPLOYMENT")
    With xmlEmploymentRequest
        .appendChild(vxmlTableElement.ownerDocument.createElement("CUSTOMERNUMBER")).Text = strCustomerNumber
        .appendChild(vxmlTableElement.ownerDocument.createElement("CUSTOMERVERSIONNUMBER")).Text = strCustomerVersionNumber
        .appendChild(vxmlTableElement.ownerDocument.createElement("EMPLOYMENTSEQUENCENUMBER")).Text = strEmploymentSequenceNumber
    End With
    Set xmlEmploymentResponse = objIDOAssist.GetData(xmlEmploymentRequest, objIClassDef.LoadEmploymentData)

    ' Append the NETMONTHLYINCOME to the main XML
    Set xmlNetMonthlyIncomeNode = xmlEmploymentResponse.selectSingleNode("NETMONTHLYINCOME")
    If Not (xmlNetMonthlyIncomeNode Is Nothing) Then
        xmlReturn.appendChild xmlNetMonthlyIncomeNode
    End If
        
    Set ICustomerEmploymentDO_FindEarnedIncomeList = xmlReturn

    Set objIClassDef = Nothing
    Set xmlClassDefDoc = Nothing
    Set objErrAssist = Nothing
    Set objIDOAssist = Nothing
    Set objXmlAssist = Nothing
    Set xmlEmploymentRequest = Nothing
    Set xmlEmploymentResponse = Nothing
    Set xmlNetMonthlyIncomeNode = Nothing
    Set xmlReturn = Nothing
    
    If Not objContext Is Nothing Then
        objContext.SetComplete
    End If
    
    Exit Function

FindEarnedIncomeListVbErr:
    
    Set objIClassDef = Nothing
    Set xmlClassDefDoc = Nothing
    Set objIDOAssist = Nothing
    Set objXmlAssist = Nothing
    Set xmlEmploymentRequest = Nothing
    Set xmlEmploymentResponse = Nothing
    Set xmlNetMonthlyIncomeNode = Nothing
    Set xmlReturn = Nothing
    
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

Private Function ICustomerEmploymentDO_FindEmploymentAndIncomeDetails(ByVal vxmlTableElement As IXMLDOMElement) As IXMLDOMNode
' header ----------------------------------------------------------------------------------
' description:
'   Get the data for all instances of the persistant data associated with
'   this data object for the values supplied
' pass:
'   vxmlxmlTableElement  xml element containing the request
' return:                xml node containing retrieved data
'------------------------------------------------------------------------------------------
On Error GoTo FindEmploymentAndIncomeDetailsVbErr

    Const strFunctionName As String = "ICustomerEmploymentDO_FindEmploymentAndIncomeDetails"

    Dim objErrAssist As New ErrAssist, _
        objXmlAssist As New XMLAssist
    
    Dim objIDOAssist As IDOAssist
    Dim objIClassDef As IomCEClassDef

    Dim nloop As Integer
    Dim xmlNode     As IXMLDOMNode, _
        xmlNodeList As IXMLDOMNodeList
    Dim xmlRootNode As IXMLDOMNode
    Dim xmlOut As New FreeThreadedDOMDocument40
    xmlOut.validateOnParse = False
    xmlOut.setProperty "NewParser", True
    Dim xmlCustomerNode   As IXMLDOMNode, _
        xmlEmploymentNode As IXMLDOMNode
    Dim strInCustomerNumber As String
    Dim strInCustomerVersionNumber As String
    Dim strInCustomerRole As String
    Dim xmlResponse As IXMLDOMNode
    Dim xmlElement As IXMLDOMElement
   
    Dim xmlEThirdParty As IXMLDOMNode, _
        xmlEDirectory  As IXMLDOMNode

    Dim objContext As ObjectContext
    Set objContext = GetObjectContext()
         
    If objContext Is Nothing Then
        Set objIDOAssist = New DOAssist
    Else
        Set objIDOAssist = objContext.CreateInstance(gstrBASE_COMPONENT & ".DOAssist")
    End If

    Set objIClassDef = New omCEClassDef

    '
    ' Get the customer list from the input XML
    '
    Set xmlNodeList = vxmlTableElement.getElementsByTagName("CUSTOMER")
    If (xmlNodeList.length < 1) Then
         ' "Invalid parameter value"
'         objErrAssist.RaiseError TypeName(Me), strFunctionName, oeInvalidParameter, "No rows found"
        ' MS Throw error not raise
         objErrAssist.ThrowError TypeName(Me), strFunctionName, oeInvalidParameter, "CUSTOMER Tag not found"
    End If
        
    ' The node that will contain the return details
    Set xmlRootNode = xmlOut.appendChild(xmlOut.createElement("EMPLOYMENTANDINCOME"))
                            
    ' Loop through each CUSTOMER in the CUSTOMERLIST in turn
    For nloop = 0 To (xmlNodeList.length - 1)
        'create temporary employment node
        Set xmlEmploymentNode = vxmlTableElement.ownerDocument.createElement("EMPLOYMENT")
        'create a customer node for the returned XML
        Set xmlCustomerNode = xmlRootNode.appendChild(xmlRootNode.ownerDocument.createElement("CUSTOMER"))

        '
        ' Retrieve the employment list for the customer
        '
        Set xmlElement = xmlNodeList.Item(nloop)
        'validate we have the CustomerNumber and CustomerVersionNumber for this customer
        strInCustomerNumber = objXmlAssist.GetTagValue(xmlElement, "CUSTOMERNUMBER")
        strInCustomerVersionNumber = objXmlAssist.GetTagValue(xmlElement, "CUSTOMERVERSIONNUMBER")
        strInCustomerRole = objXmlAssist.GetTagValue(xmlElement, "CUSTOMERROLETYPE")
        
        If Len(strInCustomerNumber) = 0 Or _
            Len(strInCustomerVersionNumber) = 0 Then
            objErrAssist.RaiseError TypeName(Me), strFunctionName, oeMissingPrimaryTag, _
            "CUSTOMERVERSION = " & strInCustomerNumber & _
            " ,CUSTOMERVERSIONNUMBER = " & strInCustomerVersionNumber & " Expected"
        End If

        'get the customer number
        Set xmlNode = xmlElement.getElementsByTagName("CUSTOMERNUMBER").Item(0)
        'add it to the temporary employment node
        xmlEmploymentNode.appendChild xmlNode.cloneNode(True)
        'get customerversionnumber
        Set xmlNode = xmlElement.getElementsByTagName("CUSTOMERVERSIONNUMBER").Item(0)
        'add it to the temporary employment node
        xmlEmploymentNode.appendChild xmlNode.cloneNode(True)
        
        'BG SYS1600 31/10/00
        Set xmlEDirectory = Nothing
        Set xmlEThirdParty = Nothing
        'BG END SYS1600 31/10/00
        
        On Error GoTo RecordNotFoundVbErr
        Set xmlEDirectory = objIDOAssist.FindListMultiple(xmlEmploymentNode, objIClassDef.LoadEmploymentDirectoryData, _
                vstrOrderByField:="CUSTOMERNUMBER,CUSTOMERVERSIONNUMBER,EMPLOYMENTSEQUENCENUMBER", vstrItemName:="EMPLOYMENT")
        Set xmlEThirdParty = objIDOAssist.FindListMultiple(xmlEmploymentNode, objIClassDef.LoadEmploymentThirdPartyData, _
                vstrOrderByField:="CUSTOMERNUMBER,CUSTOMERVERSIONNUMBER,EMPLOYMENTSEQUENCENUMBER", vstrItemName:="EMPLOYMENT")
        On Error GoTo FindEmploymentAndIncomeDetailsVbErr

        If xmlEDirectory Is Nothing Then
            Set xmlResponse = xmlEThirdParty
        ElseIf xmlEThirdParty Is Nothing Then
            Set xmlResponse = xmlEDirectory
        Else
            ' Join the thirdparty and directory xml together
            For Each xmlNode In xmlEThirdParty.childNodes
                xmlEDirectory.appendChild xmlNode.cloneNode(True)
            Next xmlNode
            Set xmlResponse = xmlEDirectory
        End If

        If Not (xmlResponse Is Nothing) Then
            For Each xmlNode In xmlResponse.selectNodes(".//EMPLOYMENT")
                If xmlNode.parentNode.nodeName <> "EMPLOYMENTLIST" Then
                    objXmlAssist.PromoteNode xmlResponse, xmlNode
                End If
            Next xmlNode
        End If
        'BG SYS1600 31/10/00 Changed code to append a clone
        If Not (xmlResponse Is Nothing) Then
            ' Add it to this customer
            xmlCustomerNode.appendChild xmlResponse.cloneNode(True)
        End If

        '
        ' Retrieve the income summary for the customer
        '
        'change node name so we can get the incomesummary
        objXmlAssist.ChangeNodeName xmlEmploymentNode, "EMPLOYMENT", "INCOMESUMMARYDETAILS"
        'get the income summary for this customer
        Set xmlResponse = Nothing
        On Error GoTo RecordNotFoundVbErr
        Set xmlResponse = objIDOAssist.GetData(xmlEmploymentNode, objIClassDef.LoadIncomeSummaryDetailsData)
        On Error GoTo FindEmploymentAndIncomeDetailsVbErr
        
        
        'BG SYS1600 31/10/00 Changed code to append a clone
        If Not (xmlResponse Is Nothing) Then
            ' Add to this customer
            xmlCustomerNode.appendChild xmlResponse.cloneNode(True)
        End If
    
        '
        ' Append the customer role type if necessary
        '
        If Len(strInCustomerRole) > 0 Then
            Set xmlNode = xmlElement.getElementsByTagName("CUSTOMERROLETYPE").Item(0)
            'add to this customer
            xmlCustomerNode.appendChild xmlNode.cloneNode(True)
        End If
    Next nloop
     
    Set ICustomerEmploymentDO_FindEmploymentAndIncomeDetails = xmlOut.documentElement

    Set objIClassDef = Nothing
    Set objErrAssist = Nothing
    Set objIDOAssist = Nothing
    Set xmlNode = Nothing
    Set xmlNodeList = Nothing
    Set xmlOut = Nothing
    Set xmlCustomerNode = Nothing
    Set xmlEmploymentNode = Nothing
    Set xmlResponse = Nothing
    Set xmlRootNode = Nothing
    Set xmlElement = Nothing
    Set objXmlAssist = Nothing
    Set xmlEThirdParty = Nothing
    Set xmlEDirectory = Nothing
    
    If Not objContext Is Nothing Then
        objContext.SetComplete
    End If
    
    Exit Function

RecordNotFoundVbErr:
    If objErrAssist.GetOmigaErrorNumber(Err.Number) = oeRecordNotFound Then
        Resume Next
    End If

FindEmploymentAndIncomeDetailsVbErr:
    
    Set objIClassDef = Nothing
    Set objIDOAssist = Nothing
    Set xmlNode = Nothing
    Set xmlNodeList = Nothing
    Set xmlOut = Nothing
    Set xmlCustomerNode = Nothing
    Set xmlEmploymentNode = Nothing
    Set xmlResponse = Nothing
    Set xmlRootNode = Nothing
    Set xmlElement = Nothing
    Set objXmlAssist = Nothing
    Set xmlEThirdParty = Nothing
    Set xmlEDirectory = Nothing
    
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

Private Function ICustomerEmploymentDO_FindEmploymentList( _
                        ByVal vxmlTableElement As IXMLDOMElement) As IXMLDOMNode
On Error GoTo FindListVbErr

    Const strFunctionName As String = "ICustomerEmploymentDO_FindEmploymentList"
    
    Dim objContext As ObjectContext
    Set objContext = GetObjectContext()
    
    Dim objIClassDef As IomCEClassDef
    Set objIClassDef = New omCEClassDef
    
    Dim objIDOAssist As IDOAssist
    
    If objContext Is Nothing Then
        Set objIDOAssist = New DOAssist
    Else
        Set objIDOAssist = objContext.CreateInstance(gstrBASE_COMPONENT & ".DOAssist")
    End If
    
    
    Set ICustomerEmploymentDO_FindEmploymentList = _
            objIDOAssist.FindListMultiple(vxmlTableElement, objIClassDef.LoadEmploymentData)

    Set objIClassDef = Nothing
    Set objIDOAssist = Nothing
    
    If Not objContext Is Nothing Then
        objContext.SetComplete
    End If
        
    Exit Function

FindListVbErr:
    
    Set objIClassDef = Nothing
    Set objIDOAssist = Nothing
    
    Dim objErrAssist As ErrAssist
    Set objErrAssist = New ErrAssist
    
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

Private Function ICustomerEmploymentDO_FindOtherIncomeList(ByVal vxmlTableElement As IXMLDOMElement) As IXMLDOMNode
' header ----------------------------------------------------------------------------------
' description:
'   Get the data for all instances of the persistant data associated with
'   this data object for the values supplied
' pass:
'   vxmlxmlTableElement  xml element containing the request
' return:                xml node containing retrieved data
'------------------------------------------------------------------------------------------
On Error GoTo FindOtherIncomeListVbErr

    Const strFunctionName As String = "ICustomerEmploymentDO_FindOtherIncomeList"

    Dim objErrAssist As New ErrAssist
    
    Dim objIDOAssist As IDOAssist
    Dim xmlClassDefDoc As FreeThreadedDOMDocument40
    Dim objIClassDef As IomCEClassDef
    
    Dim objContext As ObjectContext
    Set objContext = GetObjectContext()
         
    If objContext Is Nothing Then
        Set objIDOAssist = New DOAssist
    Else
        Set objIDOAssist = objContext.CreateInstance(gstrBASE_COMPONENT & ".DOAssist")
    End If
     
    Set objIClassDef = New omCEClassDef
    Set xmlClassDefDoc = objIClassDef.LoadUnearnedIncomeData
    Set ICustomerEmploymentDO_FindOtherIncomeList = objIDOAssist.FindList(vxmlTableElement, xmlClassDefDoc)

    Set objIClassDef = Nothing
    Set xmlClassDefDoc = Nothing
    Set objErrAssist = Nothing
    Set objIDOAssist = Nothing
    
    If Not objContext Is Nothing Then
        objContext.SetComplete
    End If
    
    Exit Function

FindOtherIncomeListVbErr:
    
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

    Set objErrAssist = Nothing
    
    '   re-raise error for business object to interpret as appropriate
    Err.Raise Err.Number, Err.Source, Err.Description

End Function

Private Function ICustomerEmploymentDO_GetAccountantDetails(ByVal vxmlTableElement As IXMLDOMElement) As IXMLDOMNode
' header ----------------------------------------------------------------------------------
' description:
'   Get the data for a single instance of the persistant data associated with
'   this data object
' pass:
'   vxmlxmlTableElement  xml element containing the request
' return:                xml node containing retrieved data
'------------------------------------------------------------------------------------------
On Error GoTo GetAccountantDetailsVbErr

    Const strFunctionName As String = "ICustomerEmploymentDO_GetAccountantDetails"
    
    Dim objErrAssist As New ErrAssist
    
    Dim objIDOAssist As IDOAssist, _
        objTPAssist As IThirdPartyAssist
    Set objTPAssist = New ThirdPartyAssist

    Dim xmlClassDefDoc As FreeThreadedDOMDocument40
    Dim objClassDef As IomCEClassDef
    Dim xmlReturn As IXMLDOMNode
    
    Dim objContext As ObjectContext
    Set objContext = GetObjectContext()
            
    If objContext Is Nothing Then
        Set objIDOAssist = New DOAssist
    Else
        Set objIDOAssist = objContext.CreateInstance(gstrBASE_COMPONENT & ".DOAssist")
    End If
    
    Set objClassDef = New omCEClassDef
    Set xmlClassDefDoc = objClassDef.LoadAccountantData
    Set xmlReturn = objIDOAssist.GetData(vxmlTableElement, xmlClassDefDoc)
    
    objTPAssist.GetThirdParty xmlReturn

    Set ICustomerEmploymentDO_GetAccountantDetails = xmlReturn

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

GetAccountantDetailsVbErr:
    
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

Private Function ICustomerEmploymentDO_GetAccountantGUID( _
                    ByVal vxmlTableElement As IXMLDOMElement) As IXMLDOMNode

                    
On Error GoTo GetDataVbErr

    Const strFunctionName As String = "ICustomerEmploymentDO_GetAccountantGUID"
            
    Dim objContext As ObjectContext
    Set objContext = GetObjectContext()
              
    Dim objIDOAssist As IDOAssist
    Dim objErrAssist As New ErrAssist
    Dim xmlTempRespNode As IXMLDOMNode
    Dim strAccountantGUID  As String
    Dim xmlClassDefDoc As FreeThreadedDOMDocument40
    Dim objIClassDef As IomCEClassDef
    
    If objContext Is Nothing Then
        Set objIDOAssist = New DOAssist
    Else
        Set objIDOAssist = objContext.CreateInstance(gstrBASE_COMPONENT & ".DOAssist")
    End If
    
    Set objIClassDef = New omCEClassDef
    Set xmlClassDefDoc = objIClassDef.LoadEmploymentAccountantGUID
    Set xmlTempRespNode = objIDOAssist.GetData(vxmlTableElement, xmlClassDefDoc)
'
'    check that the GUID was not null
'
    strAccountantGUID = xmlTempRespNode.selectSingleNode(".//ACCOUNTANTGUID").Text
    If strAccountantGUID = "" Then
      objErrAssist.ThrowError TypeName(Me), strFunctionName, oeInValidKey, "No Accountant for this Employee"
    End If
    Set ICustomerEmploymentDO_GetAccountantGUID = xmlTempRespNode.selectSingleNode("ACCOUNTANTGUID")
    Set objIClassDef = Nothing
    Set xmlClassDefDoc = Nothing
    
    If Not objContext Is Nothing Then
        objContext.SetComplete
        Set objContext = Nothing
    End If
    
    Exit Function

GetDataVbErr:
    
    Set objIClassDef = Nothing
    Set xmlClassDefDoc = Nothing
    

    
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
        
    '   re-raise error for business object to interpret as appropriate
    Err.Raise Err.Number, Err.Source, Err.Description



End Function

Private Function ICustomerEmploymentDO_GetContractDetails(ByVal vxmlTableElement As IXMLDOMElement, _
                                                          Optional ByVal vblnRetrieveEmploymentFields As Boolean = True) As IXMLDOMNode
' header ----------------------------------------------------------------------------------
' description:
'   Get the data for a single instance of the persistant data associated with
'   this data object
' pass:
'   vxmlxmlTableElement  xml element containing the request
' return:                xml node containing retrieved data
'------------------------------------------------------------------------------------------
On Error GoTo GetContractDetailsVbErr

    Const strFunctionName As String = "ICustomerEmploymentDO_GetContractDetails"
    
    Dim objErrAssist As New ErrAssist
    
    Dim objIDOAssist As IDOAssist, _
        objXmlAssist As New XMLAssist
    Dim xmlClassDefDoc As FreeThreadedDOMDocument40
    Dim objIClassDef As IomCEClassDef
    Dim xmlReturn As IXMLDOMNode
    Dim xmlContractDetailsNode As IXMLDOMNode
    Dim xmlAccountantResponse As IXMLDOMNode, _
        xmlAccountantRequest  As IXMLDOMNode, _
        xmlAccountantNameNode As IXMLDOMNode
    Dim strAccountantGUID As String
    
    Dim objContext As ObjectContext
    Set objContext = GetObjectContext()
            
    If objContext Is Nothing Then
        Set objIDOAssist = New DOAssist
    Else
        Set objIDOAssist = objContext.CreateInstance(gstrBASE_COMPONENT & ".DOAssist")
    End If
    
    Set objIClassDef = New omCEClassDef
    Set xmlClassDefDoc = objIClassDef.LoadContractDetailsProfitData()
    Set xmlReturn = objIDOAssist.FindListMultiple(vxmlTableElement, xmlClassDefDoc, vstrItemName:="CONTRACTDETAILS")
    
    If vblnRetrieveEmploymentFields Then
        ' Retrieve fields from the EMPLOYMENT table
        GetEmploymentFields xmlReturn, "CONTRACTDETAILS"
        strAccountantGUID = objXmlAssist.GetTagValue(xmlReturn, "ACCOUNTANTGUID")

        If strAccountantGUID <> "" Then
            ' Retrieve Accountant name from the accountant table
            Set xmlAccountantRequest = vxmlTableElement.ownerDocument.createElement("ACCOUNTANT")
            With xmlAccountantRequest
                .appendChild(vxmlTableElement.ownerDocument.createElement("ACCOUNTANTGUID")).Text = strAccountantGUID
            End With
            Set xmlAccountantResponse = ICustomerEmploymentDO_GetAccountantDetails(xmlAccountantRequest)
        End If
    End If

    Set xmlContractDetailsNode = xmlReturn.selectSingleNode(".//CONTRACTDETAILS")
    objXmlAssist.PromoteNode xmlReturn, xmlContractDetailsNode

    If Not ((xmlAccountantResponse Is Nothing) Or (xmlContractDetailsNode Is Nothing)) Then
        ' Append the ACCOUNTANT to the main XML
        xmlContractDetailsNode.appendChild xmlAccountantResponse
    End If

    Set ICustomerEmploymentDO_GetContractDetails = xmlReturn

    Set objIClassDef = Nothing
    Set xmlClassDefDoc = Nothing
    Set objIDOAssist = Nothing
    Set objErrAssist = Nothing
    Set xmlReturn = Nothing
    Set objXmlAssist = Nothing
    Set xmlContractDetailsNode = Nothing
    Set xmlAccountantResponse = Nothing
    Set xmlAccountantRequest = Nothing
    Set xmlAccountantNameNode = Nothing
    
    If Not objContext Is Nothing Then
        objContext.SetComplete
    End If
    
    Exit Function

GetContractDetailsVbErr:
    
    Set objIClassDef = Nothing
    Set xmlClassDefDoc = Nothing
    Set objIDOAssist = Nothing
    Set xmlReturn = Nothing
    Set objXmlAssist = Nothing
    Set xmlContractDetailsNode = Nothing
    Set xmlAccountantResponse = Nothing
    Set xmlAccountantRequest = Nothing
    Set xmlAccountantNameNode = Nothing
    
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

Private Function ICustomerEmploymentDO_GetEmployedDetails(ByVal vxmlTableElement As IXMLDOMElement, _
                                                          Optional ByVal vblnRetrieveEmploymentFields As Boolean = True) As IXMLDOMNode
' header ----------------------------------------------------------------------------------
' description:
'   Get the data for a single instance of the persistant data associated with
'   this data object
' pass:
'   vxmlxmlTableElement  xml element containing the request
' return:                xml node containing retrieved data
'------------------------------------------------------------------------------------------
On Error GoTo GetEmployedDetailsVbErr

    Const strFunctionName As String = "ICustomerEmploymentDO_GetEmployedDetails"
    
    Dim objErrAssist As New ErrAssist
    
    Dim objIDOAssist As IDOAssist, _
        objXmlAssist As New XMLAssist
    Dim xmlClassDefDoc As FreeThreadedDOMDocument40
    Dim objIClassDef As IomCEClassDef
    Dim xmlReturn As IXMLDOMNode
    Dim xmlNode     As IXMLDOMNode, _
        xmlNodeList As IXMLDOMNodeList
    Dim xmlEmployedDetailsNode As IXMLDOMNode
    Dim xmlAccountantResponse As IXMLDOMNode, _
        xmlAccountantRequest  As IXMLDOMNode, _
        xmlAccountantNameNode As IXMLDOMNode
    Dim strAccountantGUID As String
    
    Dim objContext As ObjectContext
    Set objContext = GetObjectContext()
            
    If objContext Is Nothing Then
        Set objIDOAssist = New DOAssist
    Else
        Set objIDOAssist = objContext.CreateInstance(gstrBASE_COMPONENT & ".DOAssist")
    End If
    
    Set objIClassDef = New omCEClassDef
    Set xmlClassDefDoc = objIClassDef.LoadEmployedDetailsIncomeData()
    Set xmlReturn = objIDOAssist.FindListMultiple(vxmlTableElement, xmlClassDefDoc, vstrItemName:="EMPLOYEDDETAILS")
    
    If vblnRetrieveEmploymentFields Then
        ' Retrieve fields from the EMPLOYMENT table
        GetEmploymentFields xmlReturn, "EMPLOYEDDETAILS"
        strAccountantGUID = objXmlAssist.GetTagValue(xmlReturn, "ACCOUNTANTGUID")

        If strAccountantGUID <> "" Then
            ' Retrieve Accountant name from the accountant table
            Set xmlAccountantRequest = vxmlTableElement.ownerDocument.createElement("ACCOUNTANT")
            With xmlAccountantRequest
                .appendChild(vxmlTableElement.ownerDocument.createElement("ACCOUNTANTGUID")).Text = strAccountantGUID
            End With
            Set xmlAccountantResponse = ICustomerEmploymentDO_GetAccountantDetails(xmlAccountantRequest)
        End If
    End If

    Set xmlNodeList = xmlReturn.selectNodes("EMPLOYEDDETAILSINCOME/EMPLOYEDDETAILS")
    For Each xmlNode In xmlNodeList
        objXmlAssist.PromoteNode xmlReturn, xmlNode
    Next xmlNode
    objXmlAssist.RemoveDuplicates xmlReturn, "EMPLOYMENTSEQUENCENUMBER"

    ' There should only be the one EMPLOYEDDETAILS node by this point
    objXmlAssist.GroupNodesIntoList xmlReturn.selectSingleNode("EMPLOYEDDETAILS"), "EARNEDINCOME"
    ' Make the EMPLOYEDDETAILS tag the root node
    Set xmlEmployedDetailsNode = xmlReturn.selectSingleNode("EMPLOYEDDETAILS")
    objXmlAssist.PromoteNode xmlReturn, xmlEmployedDetailsNode

    If Not ((xmlAccountantResponse Is Nothing) Or (xmlEmployedDetailsNode Is Nothing)) Then
        ' Append the ACCOUNTANT to the main XML
        xmlEmployedDetailsNode.appendChild xmlAccountantResponse
    End If

    Set ICustomerEmploymentDO_GetEmployedDetails = xmlReturn

    Set objIClassDef = Nothing
    Set xmlClassDefDoc = Nothing
    Set objIDOAssist = Nothing
    Set objErrAssist = Nothing
    Set xmlReturn = Nothing
    Set objXmlAssist = Nothing
    Set xmlNode = Nothing
    Set xmlNodeList = Nothing
    Set xmlEmployedDetailsNode = Nothing
    Set xmlAccountantResponse = Nothing
    Set xmlAccountantRequest = Nothing
    Set xmlAccountantNameNode = Nothing
    
    If Not objContext Is Nothing Then
        objContext.SetComplete
    End If
    
    Exit Function

GetEmployedDetailsVbErr:
    
    Set objIClassDef = Nothing
    Set xmlClassDefDoc = Nothing
    Set objIDOAssist = Nothing
    Set xmlReturn = Nothing
    Set objXmlAssist = Nothing
    Set xmlNode = Nothing
    Set xmlNodeList = Nothing
    Set xmlEmployedDetailsNode = Nothing
    Set xmlAccountantResponse = Nothing
    Set xmlAccountantRequest = Nothing
    Set xmlAccountantNameNode = Nothing
    
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

Private Function ICustomerEmploymentDO_GetEmploymentAndSelfEmployedDetails(ByVal vxmlTableElement As IXMLDOMElement) As IXMLDOMNode
' header ----------------------------------------------------------------------------------
' description:
'   Get the data for a single instance of the persistant data associated with
'   this data object
' pass:
'   vxmlxmlTableElement  xml element containing the request
' return:                xml node containing retrieved data
'------------------------------------------------------------------------------------------
On Error GoTo GetEmploymentAndSelfEmployedDetailsVbErr

    Const strFunctionName As String = "ICustomerEmploymentDO_GetEmploymentAndSelfEmployedDetails"
    
    Dim objErrAssist As New ErrAssist
    
    Dim objIDOAssist As IDOAssist, _
        objXmlAssist As New XMLAssist, _
        objTPAssist  As IThirdPartyAssist

    Dim xmlClassDefDoc As FreeThreadedDOMDocument40
    Dim objIClassDef As IomCEClassDef
    Dim xmlReturn  As IXMLDOMNode, _
        xmlRequest As IXMLDOMNode
    Dim xmlDetails As IXMLDOMNode, xmlDetails2 As IXMLDOMNode
    Dim objComboDO As New ComboDO
    Dim strEmploymentStatus As String
    
    Dim objContext As ObjectContext
    Set objContext = GetObjectContext()
            
    If objContext Is Nothing Then
        Set objIDOAssist = New DOAssist
        Set objTPAssist = New ThirdPartyAssist
    Else
        Set objIDOAssist = objContext.CreateInstance(gstrBASE_COMPONENT & ".DOAssist")
        Set objTPAssist = objContext.CreateInstance(App.Title & ".ThirdPartyAssist")
    End If
    
    Set objIClassDef = New omCEClassDef
    Set xmlClassDefDoc = objIClassDef.LoadEmploymentData()
    Set xmlReturn = objIDOAssist.GetData(vxmlTableElement, xmlClassDefDoc)

    ' Get third party details
    objTPAssist.GetThirdParty xmlReturn

    On Error GoTo RecordNotFoundVbErr

    objXmlAssist.ChangeNodeName vxmlTableElement, "EMPLOYMENT", "SELFEMPLOYEDDETAILS"
    Set xmlDetails = ICustomerEmploymentDO_GetSelfEmployedDetails(vxmlTableElement, vblnRetrieveEmploymentFields:=False)

    On Error GoTo GetEmploymentAndSelfEmployedDetailsVbErr

    If Not (xmlDetails Is Nothing) Then
        ' Append the details to the main XML
        xmlReturn.appendChild xmlDetails
    End If

    ' SR 08/08/2000 - SYS1385 : Append unearned income details, if any
    On Error GoTo RecordNotFoundVbErr
    objXmlAssist.ChangeNodeName vxmlTableElement, "SELFEMPLOYEDDETAILS", "UNEARNEDINCOME"
    Set xmlClassDefDoc = objIClassDef.LoadUnearnedIncomeData
    
    Set xmlDetails2 = objIDOAssist.FindList(vxmlTableElement, xmlClassDefDoc)
    
    On Error GoTo GetEmploymentAndSelfEmployedDetailsVbErr

    If Not (xmlDetails2 Is Nothing) Then
        ' Append the details to the main XML
        xmlReturn.appendChild xmlDetails2
    End If
    
    Set ICustomerEmploymentDO_GetEmploymentAndSelfEmployedDetails = xmlReturn

    Set objIClassDef = Nothing
    Set xmlClassDefDoc = Nothing
    Set objIDOAssist = Nothing
    Set objErrAssist = Nothing
    Set xmlDetails = Nothing
    Set objComboDO = Nothing
    Set xmlReturn = Nothing
    Set xmlRequest = Nothing
    Set objXmlAssist = Nothing
    Set objTPAssist = Nothing
    Set xmlDetails = Nothing
    Set xmlDetails2 = Nothing
    
    If Not objContext Is Nothing Then
        objContext.SetComplete
    End If
    
    Exit Function

RecordNotFoundVbErr:
    If objErrAssist.GetOmigaErrorNumber(Err.Number) = oeRecordNotFound Then
        ' Could not find an associated EMPLOYED/SELFEMPLOYED/CONTRACT record
        Resume Next
    End If

GetEmploymentAndSelfEmployedDetailsVbErr:
    
    Set objIClassDef = Nothing
    Set xmlClassDefDoc = Nothing
    Set objIDOAssist = Nothing
    Set xmlDetails = Nothing
    Set xmlRequest = Nothing
    Set objComboDO = Nothing
    Set xmlReturn = Nothing
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

End Function

Private Function ICustomerEmploymentDO_GetEmploymentData( _
                    ByVal vxmlTableElement As IXMLDOMElement) As IXMLDOMNode

On Error GoTo GetDataVbErr

    Const strFunctionName As String = "ICustomerEmploymentDO_GetEmploymentData"
            
    Dim objContext As ObjectContext
    Set objContext = GetObjectContext()
              
    Dim objIDOAssist As IDOAssist
              
    Dim xmlClassDefDoc As FreeThreadedDOMDocument40
    Dim objIClassDef As IomCEClassDef
    
    If objContext Is Nothing Then
        Set objIDOAssist = New DOAssist
    Else
        Set objIDOAssist = objContext.CreateInstance(gstrBASE_COMPONENT & ".DOAssist")
    End If
    
    Set objIClassDef = New omCEClassDef
    Set xmlClassDefDoc = objIClassDef.LoadEmploymentData
    Set ICustomerEmploymentDO_GetEmploymentData = objIDOAssist.GetData(vxmlTableElement, xmlClassDefDoc)

    Set objIClassDef = Nothing
    Set xmlClassDefDoc = Nothing
    
    If Not objContext Is Nothing Then
        objContext.SetComplete
        Set objContext = Nothing
    End If
    
    Exit Function

GetDataVbErr:
    
    Set objIClassDef = Nothing
    Set xmlClassDefDoc = Nothing
    
    Dim objErrAssist As ErrAssist
    Set objErrAssist = New ErrAssist
    
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
        
    '   re-raise error for business object to interpret as appropriate
    Err.Raise Err.Number, Err.Source, Err.Description
End Function

Private Function ICustomerEmploymentDO_GetEmploymentDetails(ByVal vxmlTableElement As IXMLDOMElement) As IXMLDOMNode
' header ----------------------------------------------------------------------------------
' description:
'   Get the data for a single instance of the persistant data associated with
'   this data object
' pass:
'   vxmlxmlTableElement  xml element containing the request
' return:                xml node containing retrieved data
'------------------------------------------------------------------------------------------
On Error GoTo GetEmploymentDetailsVbErr

    Const strFunctionName As String = "ICustomerEmploymentDO_GetEmploymentDetails"
    
    Dim objErrAssist As New ErrAssist
    
    Dim objIDOAssist As IDOAssist, _
        objXmlAssist As New XMLAssist, _
        objTPAssist  As IThirdPartyAssist

    Dim xmlClassDefDoc As FreeThreadedDOMDocument40
    Dim objIClassDef As IomCEClassDef
    Dim xmlReturn  As IXMLDOMNode, _
        xmlRequest As IXMLDOMNode
    Dim xmlDetails As IXMLDOMNode
    Dim objComboDO As New ComboDO
    Dim strEmploymentStatus As String
    
    Dim objContext As ObjectContext
    Set objContext = GetObjectContext()
            
    If objContext Is Nothing Then
        Set objIDOAssist = New DOAssist
        Set objTPAssist = New ThirdPartyAssist
    Else
        Set objIDOAssist = objContext.CreateInstance(gstrBASE_COMPONENT & ".DOAssist")
        Set objTPAssist = objContext.CreateInstance(App.Title & ".ThirdPartyAssist")
    End If
    
    Set objIClassDef = New omCEClassDef
    Set xmlClassDefDoc = objIClassDef.LoadEmploymentData()
    Set xmlReturn = objIDOAssist.GetData(vxmlTableElement, xmlClassDefDoc)

    ' Get third party details
    objTPAssist.GetThirdParty xmlReturn

    ' Determine the type of employment so that the appropriate CONTRACTDETAILS, SELFEMPLOYEDDETAILS or EMPLOYEDDETAILS
    ' data can be retrieved
    strEmploymentStatus = Trim$(objXmlAssist.GetTagValue(xmlReturn, "EMPLOYMENTSTATUS"))
    If strEmploymentStatus <> "" Then
        strEmploymentStatus = objComboDO.GetFirstComboValidation("EmploymentStatus", strEmploymentStatus)
    End If

    On Error GoTo RecordNotFoundVbErr
    Select Case strEmploymentStatus
        Case "C" ' Contractor
            objXmlAssist.ChangeNodeName vxmlTableElement, "EMPLOYMENT", "CONTRACTDETAILS"
            Set xmlDetails = ICustomerEmploymentDO_GetContractDetails(vxmlTableElement, _
                vblnRetrieveEmploymentFields:=False)
        Case "S" ' Self Employed
            objXmlAssist.ChangeNodeName vxmlTableElement, "EMPLOYMENT", "SELFEMPLOYEDDETAILS"
            Set xmlDetails = ICustomerEmploymentDO_GetSelfEmployedDetails(vxmlTableElement, _
                vblnRetrieveEmploymentFields:=False)
        Case "E" ' Employed
            objXmlAssist.ChangeNodeName vxmlTableElement, "EMPLOYMENT", "EMPLOYEDDETAILS"
            Set xmlDetails = ICustomerEmploymentDO_GetEmployedDetails(vxmlTableElement, _
                vblnRetrieveEmploymentFields:=False)
        Case Else
            ' Not an appropriate employment status
    End Select
    On Error GoTo GetEmploymentDetailsVbErr

    If Not (xmlDetails Is Nothing) Then
        ' Append the details to the main XML
        xmlReturn.appendChild xmlDetails
    End If

    Set ICustomerEmploymentDO_GetEmploymentDetails = xmlReturn

    Set objIClassDef = Nothing
    Set xmlClassDefDoc = Nothing
    Set objIDOAssist = Nothing
    Set objErrAssist = Nothing
    Set xmlDetails = Nothing
    Set objComboDO = Nothing
    Set xmlReturn = Nothing
    Set xmlRequest = Nothing
    Set objXmlAssist = Nothing
    Set objTPAssist = Nothing
    
    If Not objContext Is Nothing Then
        objContext.SetComplete
    End If
    
    Exit Function

RecordNotFoundVbErr:
    If objErrAssist.GetOmigaErrorNumber(Err.Number) = oeRecordNotFound Then
        ' Could not find an associated EMPLOYED/SELFEMPLOYED/CONTRACT record
        Resume Next
    End If

GetEmploymentDetailsVbErr:
    
    Set objIClassDef = Nothing
    Set xmlClassDefDoc = Nothing
    Set objIDOAssist = Nothing
    Set xmlDetails = Nothing
    Set xmlRequest = Nothing
    Set objComboDO = Nothing
    Set xmlReturn = Nothing
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
End Function

Private Function ICustomerEmploymentDO_GetNetProfitDetails(ByVal vxmlTableElement As IXMLDOMElement) As IXMLDOMNode
' header ----------------------------------------------------------------------------------
' description:
'   Get the data for a single instance of the persistant data associated with
'   this data object
' pass:
'   vxmlxmlTableElement  xml element containing the request
' return:                xml node containing retrieved data
'------------------------------------------------------------------------------------------
On Error GoTo GetNetProfitDetailsVbErr

    Const strFunctionName As String = "ICustomerEmploymentDO_GetNetProfitDetails"
    
    Dim objErrAssist As New ErrAssist
    
    Dim objIDOAssist As IDOAssist
    Dim xmlClassDefDoc As FreeThreadedDOMDocument40
    Dim objIClassDef As IomCEClassDef
    Dim objXmlAssist As New XMLAssist

    Dim strCustomerNumber           As String, _
        strCustomerVersionNumber    As String, _
        strEmploymentSequenceNumber As String
    Dim xmlEmploymentRequest    As IXMLDOMNode, _
        xmlEmploymentResponse   As IXMLDOMNode, _
        xmlNetMonthlyIncomeNode As IXMLDOMNode
    Dim xmlReturn As IXMLDOMNode
    
    Dim objContext As ObjectContext
    Set objContext = GetObjectContext()
            
    If objContext Is Nothing Then
        Set objIDOAssist = New DOAssist
    Else
        Set objIDOAssist = objContext.CreateInstance(gstrBASE_COMPONENT & ".DOAssist")
    End If

    strCustomerNumber = objXmlAssist.GetTagValue(vxmlTableElement, "CUSTOMERNUMBER")
    strCustomerVersionNumber = objXmlAssist.GetTagValue(vxmlTableElement, "CUSTOMERVERSIONNUMBER")
    strEmploymentSequenceNumber = objXmlAssist.GetTagValue(vxmlTableElement, "EMPLOYMENTSEQUENCENUMBER")
     
    Set objIClassDef = New omCEClassDef
    Set xmlClassDefDoc = objIClassDef.LoadNetProfitData()
    Set xmlReturn = objIDOAssist.GetData(vxmlTableElement, xmlClassDefDoc)

    ' Retrieve NETMONTHLYINCOME from the employment table
    Set xmlEmploymentRequest = vxmlTableElement.ownerDocument.createElement("EMPLOYMENT")
    With xmlEmploymentRequest
        .appendChild(vxmlTableElement.ownerDocument.createElement("CUSTOMERNUMBER")).Text = strCustomerNumber
        .appendChild(vxmlTableElement.ownerDocument.createElement("CUSTOMERVERSIONNUMBER")).Text = strCustomerVersionNumber
        .appendChild(vxmlTableElement.ownerDocument.createElement("EMPLOYMENTSEQUENCENUMBER")).Text = strEmploymentSequenceNumber
    End With
    Set xmlEmploymentResponse = objIDOAssist.GetData(xmlEmploymentRequest, objIClassDef.LoadEmploymentData)

    ' Append the NETMONTHLYINCOME to the main XML
    Set xmlNetMonthlyIncomeNode = xmlEmploymentResponse.selectSingleNode("NETMONTHLYINCOME")
    If Not (xmlNetMonthlyIncomeNode Is Nothing) Then
        xmlReturn.appendChild xmlNetMonthlyIncomeNode
    End If
        
    Set ICustomerEmploymentDO_GetNetProfitDetails = xmlReturn

    Set objIClassDef = Nothing
    Set xmlClassDefDoc = Nothing
    Set objIDOAssist = Nothing
    Set objErrAssist = Nothing
    Set xmlEmploymentRequest = Nothing
    Set xmlEmploymentResponse = Nothing
    Set xmlNetMonthlyIncomeNode = Nothing
    Set xmlReturn = Nothing
    Set objXmlAssist = Nothing
    
    If Not objContext Is Nothing Then
        objContext.SetComplete
    End If
    
    Exit Function

GetNetProfitDetailsVbErr:
    
    Set objIClassDef = Nothing
    Set xmlClassDefDoc = Nothing
    Set objIDOAssist = Nothing
    Set xmlEmploymentRequest = Nothing
    Set xmlEmploymentResponse = Nothing
    Set xmlNetMonthlyIncomeNode = Nothing
    Set xmlReturn = Nothing
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
End Function

Private Function ICustomerEmploymentDO_GetSelfEmployedDetails(ByVal vxmlTableElement As IXMLDOMElement, _
                                                          Optional ByVal vblnRetrieveEmploymentFields As Boolean = True) As IXMLDOMNode
' header ----------------------------------------------------------------------------------
' description:
'   Get the data for a single instance of the persistant data associated with
'   this data object
' pass:
'   vxmlxmlTableElement  xml element containing the request
' return:                xml node containing retrieved data
'------------------------------------------------------------------------------------------
On Error GoTo GetSelfEmployedDetailsVbErr

    Const strFunctionName As String = "ICustomerEmploymentDO_GetSelfEmployedDetails"
    
    Dim objErrAssist As New ErrAssist
    
    Dim objIDOAssist As IDOAssist, _
        objXmlAssist As New XMLAssist
    Dim xmlClassDefDoc As FreeThreadedDOMDocument40
    Dim objIClassDef As IomCEClassDef
    Dim xmlReturn As IXMLDOMNode
    Dim xmlSelfEmployedDetailsNode As IXMLDOMNode
    Dim xmlAccountantResponse As IXMLDOMNode, _
        xmlAccountantRequest  As IXMLDOMNode, _
        xmlAccountantNameNode As IXMLDOMNode
    Dim strAccountantGUID As String
    
    Dim objContext As ObjectContext
    Set objContext = GetObjectContext()
            
    If objContext Is Nothing Then
        Set objIDOAssist = New DOAssist
    Else
        Set objIDOAssist = objContext.CreateInstance(gstrBASE_COMPONENT & ".DOAssist")
    End If
    
    Set objIClassDef = New omCEClassDef
    Set xmlClassDefDoc = objIClassDef.LoadSelfEmployedDetailsProfitData()
    Set xmlReturn = objIDOAssist.FindListMultiple(vxmlTableElement, xmlClassDefDoc, vstrItemName:="SELFEMPLOYEDDETAILS")
    
    If vblnRetrieveEmploymentFields Then
        ' Retrieve fields from the EMPLOYMENT table
        GetEmploymentFields xmlReturn, "SELFEMPLOYEDDETAILS"
        strAccountantGUID = objXmlAssist.GetTagValue(xmlReturn, "ACCOUNTANTGUID")

        If strAccountantGUID <> "" Then
            ' Retrieve Accountant name from the accountant table
            Set xmlAccountantRequest = vxmlTableElement.ownerDocument.createElement("ACCOUNTANT")
            With xmlAccountantRequest
                .appendChild(vxmlTableElement.ownerDocument.createElement("ACCOUNTANTGUID")).Text = strAccountantGUID
            End With
            Set xmlAccountantResponse = ICustomerEmploymentDO_GetAccountantDetails(xmlAccountantRequest)
        End If
    End If

    Set xmlSelfEmployedDetailsNode = xmlReturn.selectSingleNode(".//SELFEMPLOYEDDETAILS")
    objXmlAssist.PromoteNode xmlReturn, xmlSelfEmployedDetailsNode

    If Not ((xmlAccountantResponse Is Nothing) Or (xmlSelfEmployedDetailsNode Is Nothing)) Then
        ' Append the ACCOUNTANT to the main XML
        xmlSelfEmployedDetailsNode.appendChild xmlAccountantResponse
    End If

    Set ICustomerEmploymentDO_GetSelfEmployedDetails = xmlReturn

    Set objIClassDef = Nothing
    Set xmlClassDefDoc = Nothing
    Set objIDOAssist = Nothing
    Set objErrAssist = Nothing
    Set xmlReturn = Nothing
    Set objXmlAssist = Nothing
    Set xmlSelfEmployedDetailsNode = Nothing
    Set xmlAccountantResponse = Nothing
    Set xmlAccountantRequest = Nothing
    Set xmlAccountantNameNode = Nothing
    
    If Not objContext Is Nothing Then
        objContext.SetComplete
    End If
    
    Exit Function

GetSelfEmployedDetailsVbErr:
    
    Set objIClassDef = Nothing
    Set xmlClassDefDoc = Nothing
    Set objIDOAssist = Nothing
    Set xmlReturn = Nothing
    Set objXmlAssist = Nothing
    Set xmlSelfEmployedDetailsNode = Nothing
    Set xmlAccountantResponse = Nothing
    Set xmlAccountantRequest = Nothing
    Set xmlAccountantNameNode = Nothing
    
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

Private Function ICustomerEmploymentDO_IsEmploymentStatus(ByVal venumEmploymentStatus As Integer, ByVal vstrEmploymentStatusId As String) As Boolean
' header ----------------------------------------------------------------------------------
' description:
'   Check if vstrEmploymentStatus corresponds to Self Employed,Contract,Employed or Unspecified
' pass:
'   vstrEmploymentStatus
' return:
'   IsFurtherAdvance
' Raise Errors:
'------------------------------------------------------------------------------------------
On Error GoTo IsEmploymentStatusVbErr

    Const strFunctionName As String = "ICustomerEmploymentDO_IsEmploymentStatus"

    Dim objContext As ObjectContext
    Set objContext = GetObjectContext()
    
    Dim objErrAssist As New ErrAssist
    Dim objComboDO As ComboDO

    If objContext Is Nothing Then
        Set objComboDO = New ComboDO
    Else
        Set objComboDO = objContext.CreateInstance(gstrBASE_COMPONENT & ".ComboDO")
    End If
        
    Dim strValue As String
        
    Select Case venumEmploymentStatus
        Case Employed
            strValue = "E"
        Case SelfEmployed
            strValue = "S"
        Case Contract
            strValue = "C"
        Case Unemployed
            strValue = "U"
        Case Else
            objErrAssist.RaiseError TypeName(Me), strFunctionName, _
            oeRecordNotFound
    End Select
    
    ICustomerEmploymentDO_IsEmploymentStatus = objComboDO.IsItemInValidation("EmploymentStatus", vstrEmploymentStatusId, strValue)
    
    Set objErrAssist = Nothing
    Set objComboDO = Nothing
    
    If Not objContext Is Nothing Then
        objContext.SetComplete
    End If
        
    Exit Function
        
IsEmploymentStatusVbErr:
        

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
    End If

    Set objErrAssist = Nothing
    '   re-raise error for business object to interpret as appropriate
    Err.Raise Err.Number, Err.Source, Err.Description


End Function

Private Function ICustomerEmploymentDO_IsMainSelfEmployed(ByVal vxmlElemRequest As IXMLDOMElement, _
                                Optional ByVal vxmlNode As IXMLDOMNode) As Boolean
' header ----------------------------------------------------------------------------------
' description:
'   Get the data for a single instance of the persistant data associated with
'   this data object
'   No records are trapped and just IsMainSelfEmployed is returned as FALSE
'
' pass:
'   vXMLElemRequest  xml Request element containing data to which identifies
'                   the instance of the persistant data to be retrieved
' return:
'   IsMainSelfEmployed  Boolean indicating if ANY of the customers passed in are SelfEmployed
'
'
' Raise Errors: none
'------------------------------------------------------------------------------------------
On Error GoTo IsMainSelfEmployedVbErr

    Const strFunctionName As String = "ICustomerEmploymentDO_IsMainSelfEmployed"
    
    Dim objErrAssist As New ErrAssist
    
    Dim objIDOAssist As IDOAssist
    Dim objXmlAssist As New XMLAssist
    Dim objConvertAssist As New ConvertAssist
    Dim xmlElement As IXMLDOMElement

    Dim strXML As String
    
    Dim objContext As ObjectContext
    Set objContext = GetObjectContext()
            
    If objContext Is Nothing Then
        Set objIDOAssist = New DOAssist
    Else
        Set objIDOAssist = objContext.CreateInstance(gstrBASE_COMPONENT & ".DOAssist")
    End If

    Dim strSQL As String
    Dim objIClassDef As IomCEClassDef
    Dim ERRNUMBER As Long

    Set objIClassDef = New omCEClassDef
    Set xmlElement = objIDOAssist.FindListMultiple(vxmlElemRequest, objIClassDef.LoadDataForEmployment)
    
    'DO NOT error on no record found
    
    Dim nloop As Integer
    Dim xmlNodeList As IXMLDOMNodeList
    Dim blnEmploymentStatus As Boolean
    blnEmploymentStatus = False
    Dim strTemp As String
    'get the number of employment records out
    Set xmlNodeList = xmlElement.getElementsByTagName("EMPLOYMENT")

    For nloop = 0 To (xmlNodeList.length - 1)
        
        Set xmlElement = xmlNodeList.Item(nloop)
        'check the employment status we want selfemployed
        'strTemp = objXmlAssist.GetTagValue(xmlElement, "EMPLOYMENTSTATUS")
        'MS     14/07/00    SYS0968 Need to check MAINSTATUS tag
        If objXmlAssist.GetTagValue(xmlElement, "MAINSTATUS") = "1" Then
            If objConvertAssist.CSafeBool(ICustomerEmploymentDO_IsEmploymentStatus(SelfEmployed, objXmlAssist.GetTagValue(xmlElement, "EMPLOYMENTSTATUS"))) Then
                'if any are true
                blnEmploymentStatus = True

                Exit For ' get out of this loop
            End If
        End If
                    
    Next


IsMainSelfEmployedExit:
    ICustomerEmploymentDO_IsMainSelfEmployed = blnEmploymentStatus 'return the status

    Set objIDOAssist = Nothing
    Set objXmlAssist = Nothing
    Set objErrAssist = Nothing
    Set objIClassDef = Nothing
    Set objConvertAssist = Nothing

    If Not objContext Is Nothing Then
        objContext.SetComplete
    End If
        
    Exit Function


IsMainSelfEmployedVbErr:
    
    If objErrAssist.GetOmigaErrorNumber(Err.Number) = oeRecordNotFound Then
        ERRNUMBER = Err.Number
        blnEmploymentStatus = False
        Resume IsMainSelfEmployedExit
    End If

    
    Set objIDOAssist = Nothing
    Set objXmlAssist = Nothing
    Set objIClassDef = Nothing
    Set objConvertAssist = Nothing
    
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

Private Sub ICustomerEmploymentDO_UpdateAccountantDetails(ByVal vxmlTableElement As IXMLDOMElement)
' header ----------------------------------------------------------------------------------
' description:
'   UpdateAccountantDetails a single instance of the persistant data associated with this
'   data object
' pass:
'   vxmlxmlTableElement  xml element containing the request
' return:                n/a
'------------------------------------------------------------------------------------------
On Error GoTo UpdateAccountantDetailsVbErr

    Const strFunctionName As String = "ICustomerEmploymentDO_UpdateAccountantDetails"
        
    Dim objErrAssist As New ErrAssist
    
    Dim objContext As ObjectContext
    Set objContext = GetObjectContext()
    
    Dim xmlClassDefDoc As FreeThreadedDOMDocument40
    Dim objClassDef As IomCEClassDef
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

    Set objClassDef = New omCEClassDef
    Set xmlClassDefDoc = objClassDef.LoadAccountantData

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

UpdateAccountantDetailsVbErr:
        
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

Private Sub ICustomerEmploymentDO_UpdateContractDetails(ByVal vxmlTableElement As IXMLDOMElement)
' header ----------------------------------------------------------------------------------
' description:
'   UpdateContractDetails an instance of the persistant data associated with this data object
'   for each set of data in the request
' pass:
'   vxmlxmlTableElement  xml element containing the request
' return:                n/a
'------------------------------------------------------------------------------------------
On Error GoTo UpdateContractDetailsVbErr

    Const strFunctionName As String = "ICustomerEmploymentDO_UpdateContractDetails"
   
    Dim objErrAssist As New ErrAssist, _
        objXmlAssist As New XMLAssist
    Dim xmlClassDefDoc As FreeThreadedDOMDocument40
    Dim objIClassDef As IomCEClassDef

    Dim xmlNetProfitXML As IXMLDOMNode
    Dim blnUpdate As Boolean
    
    Dim objIDOAssist As IDOAssist
    
    Dim objContext As ObjectContext
    Set objContext = GetObjectContext()
                    
    If objContext Is Nothing Then
        Set objIDOAssist = New DOAssist
    Else
        Set objIDOAssist = objContext.CreateInstance(gstrBASE_COMPONENT & ".DOAssist")
    End If
    Dim botOperationType As BOOPERATIONTYPE     'SYS2440 SA 23/7/01
    
    Set objIClassDef = New omCEClassDef
    Set xmlClassDefDoc = objIClassDef.LoadContractDetailsData()
    'SYS2440 SA 23/7/01 Problem found by Egg. {
    'Don't assume we're doing an overall update, but an update
    'of every level of the request.
    'The existing code will error if a certain element
    'of the request only has the keys passed in and tries to update
    
    '-- objIDOAssist.Update vxmlTableElement, xmlClassDefDoc
    '++ This is how it should work:
    botOperationType = objXmlAssist.DetermineOperationEx(vxmlTableElement, xmlClassDefDoc)
    Select Case botOperationType
        Case booUpdate:
            objIDOAssist.Update vxmlTableElement, xmlClassDefDoc
        Case Else
            'N/A
    End Select
    'SYS2440 end }
    
    

    '
    ' Maintain associated NETPROFIT record
    '
    blnUpdate = False
    Set xmlNetProfitXML = vxmlTableElement.selectSingleNode(".//NETPROFIT")
    If Not (xmlNetProfitXML Is Nothing) Then
        ' NETPROFIT xml exists so create or update as appropriate
        InsertEmploymentSequenceNumber xmlNetProfitXML, objXmlAssist.GetTagValue(vxmlTableElement, _
            "EMPLOYMENTSEQUENCENUMBER")

        On Error GoTo CreateError
        ICustomerEmploymentDO_CreateNetProfitDetails xmlNetProfitXML
        On Error GoTo UpdateContractDetailsVbErr

        If blnUpdate Then
            ' Error was raised whilst creating a net profit record (because record already existed)
            ICustomerEmploymentDO_UpdateNetProfitDetails xmlNetProfitXML
        End If
    End If
    
    Set objIClassDef = Nothing
    Set xmlClassDefDoc = Nothing
    Set objIDOAssist = Nothing
    Set objErrAssist = Nothing
    Set xmlNetProfitXML = Nothing
    Set objXmlAssist = Nothing

    If Not objContext Is Nothing Then
        objContext.SetComplete
    End If

    Exit Sub

CreateError:
    If objErrAssist.GetOmigaErrorNumber(Err.Number) = oeDuplicateKey Then
        blnUpdate = True
        Resume Next
    End If

UpdateContractDetailsVbErr:
    
    Set objIClassDef = Nothing
    Set xmlClassDefDoc = Nothing
    Set objIDOAssist = Nothing
    Set xmlNetProfitXML = Nothing
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

Private Sub ICustomerEmploymentDO_UpdateEarnedIncome(ByVal vxmlTableElement As IXMLDOMElement)
' header ----------------------------------------------------------------------------------
' description:
'   UpdateEarnedIncome a single instance of the persistant data associated with this
'   data object
' pass:
'   vxmlxmlTableElement  xml element containing the request
' return:                n/a
'------------------------------------------------------------------------------------------
On Error GoTo UpdateEarnedIncomeVbErr

    Const strFunctionName As String = "ICustomerEmploymentDO_UpdateEarnedIncome"
        
    Dim objErrAssist As New ErrAssist
    
    Dim objContext As ObjectContext
    Set objContext = GetObjectContext()
    
    Dim xmlClassDefDoc As FreeThreadedDOMDocument40
    Dim objIClassDef As IomCEClassDef
    Dim objIDOAssist As IDOAssist

    If objContext Is Nothing Then
        Set objIDOAssist = New DOAssist
    Else
        Set objIDOAssist = objContext.CreateInstance(gstrBASE_COMPONENT & ".DOAssist")
    End If
    
    Set objIClassDef = New omCEClassDef
    Set xmlClassDefDoc = objIClassDef.LoadEarnedIncomeData()
    objIDOAssist.Update vxmlTableElement, xmlClassDefDoc
    
    Set objIClassDef = Nothing
    Set xmlClassDefDoc = Nothing
    Set objIDOAssist = Nothing
    Set objErrAssist = Nothing
    
    If Not objContext Is Nothing Then
        objContext.SetComplete
    End If
    
    Exit Sub

UpdateEarnedIncomeVbErr:
        
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

    Set objErrAssist = Nothing
    '   re-raise error for business object to interpret as appropriate
    Err.Raise Err.Number, Err.Source, Err.Description

End Sub

Private Sub ICustomerEmploymentDO_UpdateEmployedDetails(ByVal vxmlTableElement As IXMLDOMElement)
' header ----------------------------------------------------------------------------------
' description:
'   UpdateEmployedDetails an instance of the persistant data associated with this data object
'   for each set of data in the request
' pass:
'   vxmlxmlTableElement  xml element containing the request
' return:                n/a
'------------------------------------------------------------------------------------------
On Error GoTo UpdateEmployedDetailsVbErr

    Const strFunctionName As String = "ICustomerEmploymentDO_UpdateEmployedDetails"
   
    Dim objErrAssist As New ErrAssist, _
        objXmlAssist As New XMLAssist
    Dim xmlClassDefDoc As FreeThreadedDOMDocument40
    Dim objIClassDef As IomCEClassDef

    Dim xmlEarnedIncomeXML As IXMLDOMNode
    Dim xmlNode As IXMLDOMNode
    Dim botOperationType As BOOPERATIONTYPE
    
    Dim objIDOAssist As IDOAssist
    
    Dim objContext As ObjectContext
    Set objContext = GetObjectContext()
                    
    If objContext Is Nothing Then
        Set objIDOAssist = New DOAssist
    Else
        Set objIDOAssist = objContext.CreateInstance(gstrBASE_COMPONENT & ".DOAssist")
    End If
    
    Set objIClassDef = New omCEClassDef
    Set xmlClassDefDoc = objIClassDef.LoadEmployedDetailsData()
    'SYS2440 SA 23/7/01 Problem found by Egg. {
    'Don't assume we're doing an overall update, but an update
    'of every level of the request.
    'The existing code will error if a certain element
    'of the request only has the keys passed in and tries to update
    
    '-- objIDOAssist.Update vxmlTableElement, xmlClassDefDoc
    '++ This is how it should work:
    botOperationType = objXmlAssist.DetermineOperationEx(vxmlTableElement, xmlClassDefDoc)
    Select Case botOperationType
        Case booUpdate:
            objIDOAssist.Update vxmlTableElement, xmlClassDefDoc
        Case Else
            'N/A
    End Select
    'SYS2440 end }
    
    '
    ' Maintain associated EARNEDINCOME record(s)
    '
    Set xmlClassDefDoc = objIClassDef.LoadEarnedIncomeData()
    Set xmlEarnedIncomeXML = vxmlTableElement.selectSingleNode(".//EARNEDINCOMELIST")
    If Not (xmlEarnedIncomeXML Is Nothing) Then
        ' EARNEDINCOME xml exists so create or update record(s) as appropriate
        For Each xmlNode In xmlEarnedIncomeXML.childNodes
            InsertEmploymentSequenceNumber xmlNode, objXmlAssist.GetTagValue(vxmlTableElement, _
                "EMPLOYMENTSEQUENCENUMBER")

            botOperationType = objXmlAssist.DetermineOperationEx(xmlNode, xmlClassDefDoc)

            Select Case botOperationType
                Case booCreate: ICustomerEmploymentDO_CreateEarnedIncome xmlNode
                Case booUpdate: ICustomerEmploymentDO_UpdateEarnedIncome xmlNode
                'SR 11/08/00 : SYS1433 (delete the records from UnearnedIncome)
                Case booDelete: ICustomerEmploymentDO_DeleteEarnedIncome xmlNode
                Case Else:
                    ' N/A
            End Select
        Next xmlNode
    End If
    
    'SYS1006 Maintain Employment record for SHARESOWNEDINDICATOR
    objXmlAssist.ChangeNodeName vxmlTableElement, "EMPLOYEDDETAILS", "EMPLOYMENT"
    Set xmlClassDefDoc = objIClassDef.LoadEmploymentData(omCEEmployment)
    objIDOAssist.Update vxmlTableElement, xmlClassDefDoc
    
    
    Set objIClassDef = Nothing
    Set xmlClassDefDoc = Nothing
    Set objIDOAssist = Nothing
    Set objErrAssist = Nothing
    Set xmlEarnedIncomeXML = Nothing
    Set objXmlAssist = Nothing
    Set xmlNode = Nothing

    If Not objContext Is Nothing Then
        objContext.SetComplete
    End If

    Exit Sub

UpdateEmployedDetailsVbErr:
    
    Set objIClassDef = Nothing
    Set xmlClassDefDoc = Nothing
    Set objIDOAssist = Nothing
    Set xmlEarnedIncomeXML = Nothing
    Set objXmlAssist = Nothing
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
    End If
    
    Set objErrAssist = Nothing
    
    '   re-raise error for business object to interpret as appropriate
    Err.Raise Err.Number, Err.Source, Err.Description
    
End Sub

Private Sub ICustomerEmploymentDO_UpdateEmploymentDetails(ByVal vxmlTableElement As IXMLDOMElement)
' header ----------------------------------------------------------------------------------
' description:
'   UpdateEmploymentDetails a single instance of the persistant data associated with this
'   data object
' pass:
'   vxmlxmlTableElement  xml element containing the request
' return:                n/a
'------------------------------------------------------------------------------------------
On Error GoTo UpdateEmploymentDetailsVbErr

    Const strFunctionName As String = "ICustomerEmploymentDO_UpdateEmploymentDetails"
        
    Dim objErrAssist As New ErrAssist
    
    Dim objContext As ObjectContext
    Set objContext = GetObjectContext()
    
    Dim xmlClassDefDoc As FreeThreadedDOMDocument40
    Dim objClassDef As IomCEClassDef
    Dim objIDOAssist As IDOAssist, _
        objXmlAssist As New XMLAssist, _
        objTPAssist As IThirdPartyAssist
    Set objTPAssist = New ThirdPartyAssist

    Dim xmlThirdPartyGUIDNode As IXMLDOMNode
    Dim strThirdPartyGUID As String, _
        strDirectoryGUID  As String

    Dim strEmploymentStatus As String, _
        strEmploymentSequenceNumber As String
    Dim blnUpdate As Boolean
    Dim xmlDetailsXML As IXMLDOMNode
    Dim objComboDO As New ComboDO

    If objContext Is Nothing Then
        Set objIDOAssist = New DOAssist
    Else
        Set objIDOAssist = objContext.CreateInstance(gstrBASE_COMPONENT & ".DOAssist")
    End If

    Set objClassDef = New omCEClassDef
    Set xmlClassDefDoc = objClassDef.LoadEmploymentData

    '
    ' Maintain the EMPLOYMENT (and associated THIRDPARTY record if appropriate)
    ' Note that this block of code is identical to that used in methods such as UpdateBuilder, UpdateArchitect etc.
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
    
    '
    ' Update/create any associated EMPLOYEDDETAILS, SELFEMPLOYEDDETAILS or CONTRACTDETAILS record as appropriate
    '
    strEmploymentStatus = Trim$(objXmlAssist.GetTagValue(vxmlTableElement, "EMPLOYMENTSTATUS"))
    If strEmploymentStatus <> "" Then
        strEmploymentStatus = objComboDO.GetFirstComboValidation("EmploymentStatus", strEmploymentStatus)
    End If

    strEmploymentSequenceNumber = objXmlAssist.GetTagValue(vxmlTableElement, "EMPLOYMENTSEQUENCENUMBER")
    blnUpdate = False

    Select Case strEmploymentStatus
        Case "E" ' EMPLOYEDDETAILS
            Set xmlDetailsXML = vxmlTableElement.selectSingleNode("EMPLOYEDDETAILS")
            If Not (xmlDetailsXML Is Nothing) Then
                InsertEmploymentSequenceNumber xmlDetailsXML, strEmploymentSequenceNumber
    
                On Error GoTo CreateError
                ICustomerEmploymentDO_CreateEmployedDetails xmlDetailsXML
                On Error GoTo UpdateEmploymentDetailsVbErr
    
                If blnUpdate Then
                    ' Error was raised whilst creating a EMPLOYEDDETAILS record (because record already existed)
                    ICustomerEmploymentDO_UpdateEmployedDetails xmlDetailsXML
                End If
            End If
        Case "S" ' SELFEMPLOYEDDETAILS
            Set xmlDetailsXML = vxmlTableElement.selectSingleNode("SELFEMPLOYEDDETAILS")
            If Not (xmlDetailsXML Is Nothing) Then
                InsertEmploymentSequenceNumber xmlDetailsXML, strEmploymentSequenceNumber
    
                On Error GoTo CreateError
                ICustomerEmploymentDO_CreateSelfEmployedDetails xmlDetailsXML
                On Error GoTo UpdateEmploymentDetailsVbErr
    
                If blnUpdate Then
                    ' Error was raised whilst creating a SELFEMPLOYEDDETAILS record (because record already existed)
                    ICustomerEmploymentDO_UpdateSelfEmployedDetails xmlDetailsXML
                End If
            End If
        Case "C" ' CONTRACTDETAILS
            Set xmlDetailsXML = vxmlTableElement.selectSingleNode("CONTRACTDETAILS")
            If Not (xmlDetailsXML Is Nothing) Then
                InsertEmploymentSequenceNumber xmlDetailsXML, strEmploymentSequenceNumber
    
                On Error GoTo CreateError
                ICustomerEmploymentDO_CreateContractDetails xmlDetailsXML
                On Error GoTo UpdateEmploymentDetailsVbErr
    
                If blnUpdate Then
                    ' Error was raised whilst creating a CONTRACTDETAILS record (because record already existed)
                    ICustomerEmploymentDO_UpdateContractDetails xmlDetailsXML
                End If
            End If
        Case Else
            ' No need to create/update any extra records for employment statuses of this type
    End Select

    Set objClassDef = Nothing
    Set xmlClassDefDoc = Nothing
    Set objIDOAssist = Nothing
    Set objErrAssist = Nothing
    Set objXmlAssist = Nothing
    Set objTPAssist = Nothing
    Set xmlDetailsXML = Nothing
    Set objComboDO = Nothing
    
    If Not objContext Is Nothing Then
        objContext.SetComplete
    End If
    
    Exit Sub

CreateError:
    If objErrAssist.GetOmigaErrorNumber(Err.Number) = oeDuplicateKey Then
        blnUpdate = True
        Resume Next
    End If

UpdateEmploymentDetailsVbErr:
        
    Set objClassDef = Nothing
    Set xmlClassDefDoc = Nothing
    Set objIDOAssist = Nothing
    Set objXmlAssist = Nothing
    Set objTPAssist = Nothing
    Set xmlDetailsXML = Nothing
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
    End If

    Set objErrAssist = Nothing
    '   re-raise error for business object to interpret as appropriate
    Err.Raise Err.Number, Err.Source, Err.Description

End Sub

Private Sub ICustomerEmploymentDO_UpdateNetProfitDetails(ByVal vxmlTableElement As IXMLDOMElement)
' header ----------------------------------------------------------------------------------
' description:
'   UpdateNetProfitDetails a single instance of the persistant data associated with this
'   data object
' pass:
'   vxmlxmlTableElement  xml element containing the request
' return:                n/a
'------------------------------------------------------------------------------------------
On Error GoTo UpdateNetProfitDetailsVbErr

    Const strFunctionName As String = "ICustomerEmploymentDO_UpdateNetProfitDetails"
    'SYS2440 Get XMLAssist too
    Dim objErrAssist As New ErrAssist, _
        objXmlAssist As New XMLAssist
        
    Dim objContext As ObjectContext
    Set objContext = GetObjectContext()
    
    Dim xmlClassDefDoc As FreeThreadedDOMDocument40
    Dim objIClassDef As IomCEClassDef
    Dim objIDOAssist As IDOAssist
    Dim botOperationType As BOOPERATIONTYPE     'SYS2440 SA 23/7/01
   
    If objContext Is Nothing Then
        Set objIDOAssist = New DOAssist
    Else
        Set objIDOAssist = objContext.CreateInstance(gstrBASE_COMPONENT & ".DOAssist")
    End If
    
    Set objIClassDef = New omCEClassDef
    Set xmlClassDefDoc = objIClassDef.LoadNetProfitData()
    'SYS2440 SA 23/7/01 Problem found by Egg. {
    'Don't assume we're doing an overall update, but an update
    'of every level of the request.
    'The existing code will error if a certain element
    'of the request only has the keys passed in and tries to update
    
    '-- objIDOAssist.Update vxmlTableElement, xmlClassDefDoc
    '++ This is how it should work:
    botOperationType = objXmlAssist.DetermineOperationEx(vxmlTableElement, xmlClassDefDoc)
    Select Case botOperationType
        Case booUpdate:
            objIDOAssist.Update vxmlTableElement, xmlClassDefDoc
        Case Else
            'N/A
    End Select
    'SYS2440 end }
    
    Set objIClassDef = Nothing
    Set xmlClassDefDoc = Nothing
    Set objIDOAssist = Nothing
    Set objErrAssist = Nothing
    
    If Not objContext Is Nothing Then
        objContext.SetComplete
    End If
    
    Exit Sub

UpdateNetProfitDetailsVbErr:
        
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

    Set objErrAssist = Nothing
    '   re-raise error for business object to interpret as appropriate
    Err.Raise Err.Number, Err.Source, Err.Description

End Sub

Private Sub ICustomerEmploymentDO_UpdateOtherIncome(ByVal vxmlTableElement As IXMLDOMElement)
' header ----------------------------------------------------------------------------------
' description:
'   UpdateOtherIncome a single instance of the persistant data associated with this
'   data object
' pass:
'   vxmlxmlTableElement  xml element containing the request
' return:                n/a
'------------------------------------------------------------------------------------------
On Error GoTo UpdateOtherIncomeVbErr

    Const strFunctionName As String = "ICustomerEmploymentDO_UpdateOtherIncome"
        
    Dim objErrAssist As New ErrAssist
    
    Dim objContext As ObjectContext
    Set objContext = GetObjectContext()
    
    Dim xmlClassDefDoc As FreeThreadedDOMDocument40
    Dim objIClassDef As IomCEClassDef
    Dim objIDOAssist As IDOAssist

    If objContext Is Nothing Then
        Set objIDOAssist = New DOAssist
    Else
        Set objIDOAssist = objContext.CreateInstance(gstrBASE_COMPONENT & ".DOAssist")
    End If
    
    Set objIClassDef = New omCEClassDef
    Set xmlClassDefDoc = objIClassDef.LoadUnearnedIncomeData()
    objIDOAssist.Update vxmlTableElement, xmlClassDefDoc
    
    Set objIClassDef = Nothing
    Set xmlClassDefDoc = Nothing
    Set objIDOAssist = Nothing
    Set objErrAssist = Nothing
    
    If Not objContext Is Nothing Then
        objContext.SetComplete
    End If
    
    Exit Sub

UpdateOtherIncomeVbErr:
        
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

    Set objErrAssist = Nothing
    '   re-raise error for business object to interpret as appropriate
    Err.Raise Err.Number, Err.Source, Err.Description

End Sub

Private Sub ICustomerEmploymentDO_UpdateSelfEmployedDetails(ByVal vxmlTableElement As IXMLDOMElement)
' header ----------------------------------------------------------------------------------
' description:
'   UpdateSelfEmployedDetails an instance of the persistant data associated with this data object
'   for each set of data in the request
' pass:
'   vxmlxmlTableElement  xml element containing the request
' return:                n/a
'------------------------------------------------------------------------------------------
On Error GoTo UpdateSelfEmployedDetailsVbErr

    Const strFunctionName As String = "ICustomerEmploymentDO_UpdateSelfEmployedDetails"
   
    Dim objErrAssist As New ErrAssist, _
        objXmlAssist As New XMLAssist
    Dim xmlClassDefDoc As FreeThreadedDOMDocument40
    Dim objIClassDef As IomCEClassDef

    Dim xmlNetProfitXML As IXMLDOMNode
    Dim blnUpdate As Boolean
    Dim botOperationType As BOOPERATIONTYPE     'SYS2440 SA 23/7/01
   
    
    Dim objIDOAssist As IDOAssist
    
    Dim objContext As ObjectContext
    Set objContext = GetObjectContext()
                    
    If objContext Is Nothing Then
        Set objIDOAssist = New DOAssist
    Else
        Set objIDOAssist = objContext.CreateInstance(gstrBASE_COMPONENT & ".DOAssist")
    End If
    
    Set objIClassDef = New omCEClassDef
    Set xmlClassDefDoc = objIClassDef.LoadSelfEmployedDetailsData()
    
    'SYS2440 SA 23/7/01 Problem found by Egg. {
    'Don't assume we're doing an overall update, but an update
    'of every level of the request.
    'The existing code will error if a certain element
    'of the request only has the keys passed in and tries to update
    
    '-- objIDOAssist.Update vxmlTableElement, xmlClassDefDoc
    '++ This is how it should work:
    botOperationType = objXmlAssist.DetermineOperationEx(vxmlTableElement, xmlClassDefDoc)
    Select Case botOperationType
        Case booUpdate:
            objIDOAssist.Update vxmlTableElement, xmlClassDefDoc
        Case Else
            'N/A
    End Select
    'SYS2440 end }
    
    '
    ' Maintain associated NETPROFIT record
    '
    blnUpdate = False
    Set xmlNetProfitXML = vxmlTableElement.selectSingleNode(".//NETPROFIT")
    If Not (xmlNetProfitXML Is Nothing) Then
        ' NETPROFIT xml exists so create or update as appropriate
        InsertEmploymentSequenceNumber xmlNetProfitXML, objXmlAssist.GetTagValue(vxmlTableElement, _
            "EMPLOYMENTSEQUENCENUMBER")

        On Error GoTo CreateError
        ICustomerEmploymentDO_CreateNetProfitDetails xmlNetProfitXML
        On Error GoTo UpdateSelfEmployedDetailsVbErr

        If blnUpdate Then
            ' Error was raised whilst creating a net profit record (because record already existed)
            ICustomerEmploymentDO_UpdateNetProfitDetails xmlNetProfitXML
        End If
    End If
    
    Set objIClassDef = Nothing
    Set xmlClassDefDoc = Nothing
    Set objIDOAssist = Nothing
    Set objErrAssist = Nothing
    Set xmlNetProfitXML = Nothing
    Set objXmlAssist = Nothing

    If Not objContext Is Nothing Then
        objContext.SetComplete
    End If

    Exit Sub

CreateError:
    If objErrAssist.GetOmigaErrorNumber(Err.Number) = oeDuplicateKey Then
        blnUpdate = True
        Resume Next
    End If

UpdateSelfEmployedDetailsVbErr:
    
    Set objIClassDef = Nothing
    Set xmlClassDefDoc = Nothing
    Set objIDOAssist = Nothing
    Set xmlNetProfitXML = Nothing
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

Private Sub ICustomerEmploymentDO_UpdateTaxDetails(ByVal vxmlTableElement As IXMLDOMElement)
' header ----------------------------------------------------------------------------------
' description:
'   UpdateTaxDetails a single instance of the persistant data associated with this
'   data object
' pass:
'   vxmlxmlTableElement  xml element containing the request
' return:                n/a
'------------------------------------------------------------------------------------------
On Error GoTo UpdateTaxDetailsVbErr

    Const strFunctionName As String = "ICustomerEmploymentDO_UpdateTaxDetails"
        
    Dim objErrAssist As New ErrAssist
    
    Dim objContext As ObjectContext
    Set objContext = GetObjectContext()
    
    Dim xmlClassDefDoc As FreeThreadedDOMDocument40
    Dim objIClassDef As IomCEClassDef
    Dim objIDOAssist As IDOAssist

    If objContext Is Nothing Then
        Set objIDOAssist = New DOAssist
    Else
        Set objIDOAssist = objContext.CreateInstance(gstrBASE_COMPONENT & ".DOAssist")
    End If
    
    Set objIClassDef = New omCEClassDef
    Set xmlClassDefDoc = objIClassDef.LoadIncomeSummaryData()
    objIDOAssist.Update vxmlTableElement, xmlClassDefDoc
    
    Set objIClassDef = Nothing
    Set xmlClassDefDoc = Nothing
    Set objIDOAssist = Nothing
    Set objErrAssist = Nothing
    
    If Not objContext Is Nothing Then
        objContext.SetComplete
    End If
    
    Exit Sub

UpdateTaxDetailsVbErr:
        
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

    Set objErrAssist = Nothing
    '   re-raise error for business object to interpret as appropriate
    Err.Raise Err.Number, Err.Source, Err.Description

End Sub

Private Function ICustomerEmploymentDO_GetTaxDetails(ByVal vxmlTableElement As IXMLDOMElement) As IXMLDOMNode
' header ----------------------------------------------------------------------------------
' description:
'   Get the data for a single instance of the persistant data associated with
'   this data object
' pass:
'   vxmlTableElement  xml element containing the request
' return:                xml node containing retrieved data
'------------------------------------------------------------------------------------------
On Error GoTo GetTaxDetailsVbErr

    Const strFunctionName As String = "ICustomerEmploymentDO_GetTaxDetails"
    
                            
    Dim objContext As ObjectContext
    Set objContext = GetObjectContext()
    
    Dim objErrAssist As ErrAssist
    Set objErrAssist = New ErrAssist
    
    Dim xmlClassDefDoc As FreeThreadedDOMDocument40
    Dim objIDOAssist As IDOAssist
    Dim objIClassDef As IomCEClassDef
    
    
    Set objIDOAssist = New DOAssist

    
    Set objIClassDef = New omCEClassDef
    Set xmlClassDefDoc = objIClassDef.LoadIncomeSummaryData
    Set ICustomerEmploymentDO_GetTaxDetails = objIDOAssist.GetData(vxmlTableElement, xmlClassDefDoc)
    
    objContext.SetComplete
    
    Set objIClassDef = Nothing
    Set xmlClassDefDoc = Nothing
    Set objIDOAssist = Nothing
    Set objContext = Nothing

    
    
    Exit Function

GetTaxDetailsVbErr:
    
    
    Set objIClassDef = Nothing
    Set xmlClassDefDoc = Nothing
    Set objIDOAssist = Nothing
    
    objErrAssist.AddToErrSource (strFunctionName)
    
    If objErrAssist.IsSystemError = True Then
        objErrAssist.LogError TypeName(Me), strFunctionName, Err.Number, Err.Description
    End If
    
    If objErrAssist.IsSystemError Then
        objContext.SetAbort
    Else
        objContext.SetComplete
    End If
    
    Set objContext = Nothing
   
    '   re-raise error for business object to interpret as appropriate
    Err.Raise Err.Number, Err.Source, Err.Description
End Function
'++BMIDS00246 New method added.
Private Function ICustomerEmploymentDO_GetEmploymentStatusList(ByVal vxmlElemRequest As IXMLDOMElement) As String
' header ----------------------------------------------------------------------------------
' description:
'   Get the data for a single instance of the persistant data associated with
'   this data object
'   No records are trapped and the main employment status for each applicant is returned as string
'
' pass:
'   vXMLElemRequest  xml Request element containing data to which identifies
'                   the instance of the persistant data to be retrieved
' return:
'   GetEmploymentStatus  String of comma separated EmploymentStatus values for main employments of all applicants
'
'
' Raise Errors: none
'------------------------------------------------------------------------------------------
On Error GoTo GetEmploymentStatusListVbErr

    Const strFunctionName As String = "ICustomerEmploymentDO_GetEmploymentStatusList"
    
    Dim objErrAssist As New ErrAssist
    
    Dim objIDOAssist As IDOAssist
    Dim objXmlAssist As New XMLAssist
    Dim objConvertAssist As New ConvertAssist
    Dim xmlElement As IXMLDOMElement

    Dim strXML As String
    
    Dim objContext As ObjectContext
    Set objContext = GetObjectContext()
            
    If objContext Is Nothing Then
        Set objIDOAssist = New DOAssist
    Else
        Set objIDOAssist = objContext.CreateInstance(gstrBASE_COMPONENT & ".DOAssist")
    End If

    Dim objIClassDef As IomCEClassDef
    Dim ERRNUMBER As Long

    Set objIClassDef = New omCEClassDef
    Set xmlElement = objIDOAssist.FindListMultiple(vxmlElemRequest, objIClassDef.LoadDataForEmployment)
    
    'DO NOT error on no record found
    
    Dim nloop As Integer
    Dim xmlNodeList As IXMLDOMNodeList
    Dim strEmploymentStatus As String
    strEmploymentStatus = ""
    
    Dim strTemp As String
    'get the number of employment records out
    Set xmlNodeList = xmlElement.getElementsByTagName("EMPLOYMENT")

    For nloop = 0 To (xmlNodeList.length - 1)
        
        Set xmlElement = xmlNodeList.Item(nloop)
        'Only interested in main employments
        If objXmlAssist.GetTagValue(xmlElement, "MAINSTATUS") = "1" Then
            'Build up string to return
            'BMIDS00662 - allow for single apps with previous employment history!
            '-- If nloop = 0 Then
            If Len(strEmploymentStatus) = 0 Then
                'first time in
                strEmploymentStatus = objXmlAssist.GetTagValue(xmlElement, "EMPLOYMENTSTATUS")
            Else
                strEmploymentStatus = strEmploymentStatus & ", " & objXmlAssist.GetTagValue(xmlElement, "EMPLOYMENTSTATUS")
            End If
        End If
    Next
    
    

GetEmploymentStatusListExit:
    
    ICustomerEmploymentDO_GetEmploymentStatusList = strEmploymentStatus 'return the string of employment status'

    Set objIDOAssist = Nothing
    Set objXmlAssist = Nothing
    Set objErrAssist = Nothing
    Set objIClassDef = Nothing
    Set objConvertAssist = Nothing

    If Not objContext Is Nothing Then
        objContext.SetComplete
    End If
        
    Exit Function


GetEmploymentStatusListVbErr:
    
    If objErrAssist.GetOmigaErrorNumber(Err.Number) = oeRecordNotFound Then
        ERRNUMBER = Err.Number
        Resume GetEmploymentStatusListExit
    End If

    
    Set objIDOAssist = Nothing
    Set objXmlAssist = Nothing
    Set objIClassDef = Nothing
    Set objConvertAssist = Nothing
    
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
'++BMIDS00246 }


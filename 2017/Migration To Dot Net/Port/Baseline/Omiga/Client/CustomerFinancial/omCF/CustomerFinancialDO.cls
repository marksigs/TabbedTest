VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 3  'UsesTransaction
END
Attribute VB_Name = "CustomerFinancialDO"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
'Workfile:      Omiga4DO.cls
'Copyright:     Copyright © 1999 Marlborough Stirling

'Description:   Code template for omiga4 Data Object
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
'MCS    17/08/99 Created
'MCS    30/08/99 Update following code review
'MCS    01/09/99 Optional timimg output added
'IW     02/05/00 SYS0601 - Removed PolicyCreation
'IW     05/05/00 SYS0138 - Added ApplicationContract Link Table
'MC     08/05/00 SYS0679 - Use MortgageAccountThirdParty view to retrieve mortgage accounts
'                           for arrears details screen DC131
'MH     08/05/00 SYS0255 - Fixed FindLoansAndLiabilitiesWithNoLender -
'                           it was not appending output nodes correctly.
'APS    01/06/00 SYS0800 - Removed Record Not Found processing from FindRegularOutgoingsList
'MC     08/06/00 SYS0866 - Standardise behaviour of Balance & Payment fields
'MH     08/05/00 SYS0255 - Fixed FindLoansAndLiabilitiesWithNoLender -
'                          DomDocument confusion. See instream
'SR     23/06/00 Modified DeleteLiability
'MS     27/06/99 Re-raise error correctly
'JLD    27/06/00 cope with no LOANSLIABILITIES in FindLiabilityList
'BG     27/06/00 SYS0928 - handles empty primary tags
'PSC    28/06/00 SYS1012 - Add new method GetLoansNotRedeemedTotals
'SR     07/07/00 SYS0892  Modified method 'FindMortgageLoanList' - add a optional input parameter
'                to sort the XML in descending order of StartDate
'SR     13/07/00 SYS1201 - Modified mehtod UpdateMortgageAccount - ignore, if IndemnityInsurance
'                          is not found in the request.
'JLD    18/08/00 SYS1437 - added ORDER BY to FindMortgageAccountList so the order can be predictable
'BG     03/10/00 SYS0419 - Use LoadMortgageAccountArrearsData rather than
'                          LoadMortgageAccountThirdPartyData in FindMortgageListForArrears method.
'LD     07/11/00           Explicity close database connections
'LD     07/11/00           Explicity close recordsets
'ASt    09/03/01           CC012 GetMortgageLoanBalance now compatible with SQL Server and Oracle.
'GD     14/03/01           SYS2047 Added GetMortgageAccount
'DRC    22/03/01 SYS1932   Added GetAccountDetails method
'LD     11/06/01 SYS2367 - SQL Server Port - Length must be specified in calls to CreateParameter
'DRC    14/06/01 SYS0999 - Method GetMortgageLenderDirectoryData added
'LD    19/06/01  SYS2386 All projects to use guidassist.bas rather than guidassist.cls
'GD     27/07/01 SYS2514   Reset List For Each Applicant - XML Parser Issue
'DM     09/09/01 SYS2572 Added Generic sql processing.
'JR     05/10/01 Omiplus24, Modified FindLiabilityList and added new private method AppendContactTelephoneData
'AD     09/10/01 Removed the MoveFirst as it's a ForwardOnly cursor. SYS2718.
'JR     29/10/01 SYS2860   Comment out ref. to Omiplus24 telephone number changes until further notice
'JR     03/12/01 SYS2459   Re-apply Omiplus24 changes to FindLiabilityList & added AppendContactTelephoneData
'JLD    21/01/02 SYS3844   when calculating liability totals in FindLiabilityList only sum unique records.
'MH     24/04/06 MAR1448    Remove Indemnity Insurance xml when record not present in database
'------------------------------------------------------------------------------------------------
'Epsom history
'
'Prog   Date        Description
'GHun   26/10/2006  EP2_9 Changed GetMtgAcctDataForFurtherAdv to removed BM specific code and return LastValuationDate
'AShaw  29/12/2006  EP2_56 Changed ICustomerFinancialDO_GetPortedStepAndPeriodFromMortgageAccount to add new param.
'GHun   30/01/2007  EP2_954 Added new function GetAdditionalBorrowingDataForProjectionCalc
'------------------------------------------------------------------------------------------------

Option Explicit

Implements ICustomerFinancialDO
Implements ObjectControl

Private m_objContext As ObjectContext
Private m_objErrAssist As ErrAssist
Private m_objXmlAssist As XMLAssist             ' Remove if not required
Private m_objIDOAssist As IDOAssist


'-------------------------------------------------------------------------
'BMids History:
'
'Prog   Date        Description
'PSC    17/07/02    BMIDS00006  Added CreateAccountRelationships
'                                    DeleteAccountRelationships
'                                    UpdateAccountRelationships
'                                    CreateSpecialFeature
'                                    UpdateSpecialFeature
'                                    DeleteSpecialFeature
'                                    FindSpecialFeatureList
'                                    GetMortgageAccountDetails
'                               Amend CreateMortgageAccount to use new structure
'                                    GetLoansNotRedeemedTotals to use new structure
'                                    FindMortgageAccountList to use new structure
'                                    UpdateMortgageAccount to use new structure
'                                    FindArrearsHistoryList to use new structure
'PSC    13/08/2002  BMIDS00006  Amend GetMortgageAccountDetails not to rename MAADDRESSGUID
'GHun   24/07/2002  BMIDS00190  DCWP3 BM076 Changed CreateBankruptcyHistory,
'                                                  CreateCCJHistory,
'                                                  CreateRegularOutgoings,
'                                                  CreateLiability,
'                                                  CreateArrearsHistory,
'                                                  DeleteBankruptcyHistory,
'                                                  DeleteCCJHistory,
'                                                  DeleteRegularOutgoings,
'                                                  DeleteArrearsHistory,
'                                                  DeleteLiability,
'                                                  FindBankruptcyHistoryList,
'                                                  FindCCJHistoryList,
'                                                  FindRegularOutgoingsList,
'                                                  FindArrearsHistoryList,
'                                                  FindLiabilityList,
'                                                  UpdateBankruptcyHistory,
'                                                  UpdateCCJHistory,
'                                                  UpdateRegularOutgoings,
'                                                  UpdateArrearsHistory,
'                                                  UpdateLiability,
'                                                  FindMortgageListForArrears,
'                                           Added  FindLoansListForArrears
'GHun   29/08/2002  BMIDS00383  Changed FindLiabilityList to handle not receiving data back
'MV     02/09/2002  BMIDS00353  CMWP4 - BM087 ; Added GetMtgAcctDataForFurtherAdv
'GHun   04/09/2002  BMIDS00403  UpdateRegularOutgoings not updating customers
'GHun   05/09/2002  BMIDS00405  FindMortgageAccountList - remove duplicates from list
'GHun   05/09/2002  BMIDS00409  Changed FindLoansToBePorted to avoid infinite loop
'MDC    10/09/2002  BMIDS00336  CCWP1 BM062 - Add GetMortgageAccountLoanTotals
'MDC    27/09/2002  BMIDS00518  Fix GetMortgageAccountDetails error
'MDC    30/09/2002  BMIDS00510  Security Address
'MV     03/10/2002  BMIDS00430  Modified ICustomerFinancialDO_GetMtgAcctDataForFurtherAdv
'MV     23/10/2002  BMIDS00430  Rewritten ICustomerFinancialDO_GetMtgAcctDataForFurtherAdv
'MDC    12/11/2002  BMIDS00919  Validate AddressGuid in GetMortgageAccountDetails
'GHun   14/11/2002  BMIDS00444  Added FindRemortgageAccountAddress
'GHun   19/11/2002  BMIDS00973  Amend DeleteLiability to also delete associated ArrearsHistory records
'MV     04/03/2003  BM0379      Amended ICustomerFinancialDO_FindAccountRelationshipList
'MV     09/04/2003  BM0502      Amended ICustomerFinancialDO_FindAccountRelationshipList
'LDM    15/05/2003  BM0492      Amended GetMortgageLoanBalance to return outstanding balance of total loans not redeemed if parameter requests it
'GHun   01/07/2003  BMIDS591    CC040 Amended GetMtgAcctDataForFurtherAdv to only make one stored proc call and to pass an extra parameter
'SR     25/05/2004  BMIDS772    New methods ICustomerFinancialDO_GetCreditHistorySummaryView and ICustomerFinancialDO_GetFinancialSummaryView
'SR     26/05/2004  BMIDS772    Modified methods ICustomerFinancialDO_GetFinancialSummaryView,
'                                                ICustomerFinancialDO_GetCreditHistorySummaryView
' RF    19/07/04    BMIDS788    Improve performance of ImportAccountsIntoApplication
'SR     03/09/2004  BMIDS815
'GHun   07/09/2004  BMIDS815    Amended GetPortedStepAndPeriodFromMortgageAccount
'JD     06/10/2004  BMIDS896    DeleteLiability now uses stored procedure
'------------------------------------------------------------------------------------------------------------------------------------
'MARS History:
'
'Prog   Date        Description
'GHun   12/10/2005  MAR46       Changed GetPortedStepAndPeriodFromMortgageAccount
'------------------------------------------------------------------------------------------------------------------------------------

Private Function ICustomerFinancialDO_GetAccountDetails(ByVal vxmlTableElement As IXMLDOMElement) As IXMLDOMNode
' header ----------------------------------------------------------------------------------
' description:
'   Get the data for a single instance of the persistant data associated with
'   this data object
' pass:
'   vxmlxmlTableElement  xml element containing the request
' return:                xml node containing retrieved data
'------------------------------------------------------------------------------------------
On Error GoTo GetAccountDetailsVbErr

    Const strFunctionName As String = "ICustomerFinancialDO_GetAccountDetails"
    
    Dim objErrAssist As New ErrAssist
    
    Dim objIDOAssist As IDOAssist, _
        objTPAssist As IThirdPartyAssist
    Set objTPAssist = New ThirdPartyAssist

    Dim xmlClassDefDoc As FreeThreadedDOMDocument40
    Dim objClassDef As ICFClassDef
    Dim xmlReturn As IXMLDOMNode
    
    Dim objContext As ObjectContext
    Set objContext = GetObjectContext()
            
    If objContext Is Nothing Then
        Set objIDOAssist = New DOAssist
    Else
        Set objIDOAssist = objContext.CreateInstance(gstrBASE_COMPONENT & ".DOAssist")
    End If
    
    Set objClassDef = New CFClassDef
    Set xmlClassDefDoc = objClassDef.LoadAccountData
    Set xmlReturn = objIDOAssist.GetData(vxmlTableElement, xmlClassDefDoc)
    
    objTPAssist.GetThirdParty xmlReturn

    Set ICustomerFinancialDO_GetAccountDetails = xmlReturn

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

GetAccountDetailsVbErr:
    
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

Private Function MuraliGetLoansNotRedeemedTotals(ByVal vxmlTableElement As IXMLDOMElement) As IXMLDOMNode

    On Error GoTo GetLoansNotRedeemedTotalsVbErr

    Const strFunctionName As String = "MuraliGetLoansNotRedeemedTotals"
    
    Dim xmlCustomerNodeList As IXMLDOMNodeList
    Dim xmlCustomer As IXMLDOMElement
    Dim xmlClassDefDoc As FreeThreadedDOMDocument40
    
    Dim adoConnection As New ADODB.Connection
    Dim adoCommand As New ADODB.Command
    Dim rstThisRecSet As New ADODB.Recordset
    Dim adoParameter As ADODB.Parameter
    
    Dim objIADOAssist As IADOAssist
    Dim objIDOAssist As IDOAssist
    Dim objIomCFClassDef As ICFClassDef
    
    Dim strCustomerNo As String
    Dim strCustomerVersionNo As String
    
    If m_objContext Is Nothing Then
        Set objIADOAssist = New ADOAssist
        Set objIDOAssist = New DOAssist
    Else
        Set objIADOAssist = m_objContext.CreateInstance(gstrBASE_COMPONENT & ".ADOAssist")
        Set objIDOAssist = m_objContext.CreateInstance(gstrBASE_COMPONENT & ".DOAssist")
    End If
    
    Set xmlCustomerNodeList = vxmlTableElement.selectNodes(".//CUSTOMER")
    
    If xmlCustomerNodeList.length = 0 Then
        m_objErrAssist.ThrowError TypeName(Me), _
                                  strFunctionName, _
                                  oeMissingPrimaryTag, _
                                  "CUSTOMERVERSION tag missing"
    End If
    
    'Stage1:
        
    ' Set MonthlyRepayment and Outstanding Balance from all MortgageLoan Records
    'strApplicationNumber = Trim$(objXmlAssist.GetTagValue(vxmlRequest, "APPLICATIONNUMBER"))
    'strApplicationFactFindNumber = Trim$(objXmlAssist.GetTagValue(vxmlRequest, "APPLICATIONFACTFINDNUMBER"))
    
    'If Len(strApplicationNumber) = 0 Or Len(strApplicationFactFindNumber) = 0 Then
    '    m_objErrAssist.ThrowError TypeName(Me), strFunctionName, oeInvalidParameter, _
    '                    "ApplicationNumber & ApplicationFactFindNumber must be specified"
    'End If
   '
   ' #If GENERIC_SQL Then
   '     Dim eDbEngineType As DBENGINETYPE
    '    eDbEngineType = objIADOAssist.GetDBEngine
'        If eDbEngineType = SQLServer Then
'            strSql = "USP_GETMORTAGELOANTOTALS"
'        End If
'    #End If
'
'    With adoConnection
'        .ConnectionString = objIADOAssist.GetConnStr
'        #If Not GENERIC_SQL Then
'            .CursorLocation = adUseClient
'        #End If
'        .Open
'    End With
'
'    Set adoCommand = New ADODB.Command
'    Set adoCommand.ActiveConnection = adoConnection
'    adoCommand.CommandType = adCmdStoredProc
'    adoCommand.CommandText = strSql
'
'    ' Application Number
'    Set adoParameter = adoCommand.CreateParameter("APPLICATIONNUMBER", adBSTR, adParamInput, Len(strApplicationNumber))
'    adoParameter.Attributes = adParamNullable
'    adoCommand.Parameters.Append adoParameter
'    If Len(strApplicationNumber) > 0 Then
'        adoCommand(0) = strApplicationNumber
'    Else
'        adoCommand(0) = Null
'    End If
'
'    ' ApplictionFactFindNumber
'    Set adoParameter = adoCommand.CreateParameter("APPLICATIONFACTFINDNUMBER", adInteger, adParamInput, Len(strApplicationFactFindNumber))
'    adoParameter.Attributes = adParamNullable
'    adoCommand.Parameters.Append adoParameter
'    If Len(strApplicationFactFindNumber) > 0 Then
'        adoCommand(1) = strApplicationFactFindNumber
'    Else
'        adoCommand(1) = Null
'    End If
'
'    Set adoTotalMonthlyRepayment = adoCommand.CreateParameter("TOTALMONTHLYREPAYMENT", adDouble, adParamOutput)
'    adoCommand.Parameters.Append adoTotalMonthlyRepayment
'
'    Set adoTotalOutstandingBalance = adoCommand.CreateParameter("TOTALOUTSTANDINGBALANCE", adDouble, adParamOutput)
'    adoCommand.Parameters.Append adoTotalOutstandingBalance
'
'    adoCommand.Execute
'
'    dblMonthlyRepayment = objConvAssist.CSafeDbl(adoTotalMonthlyRepayment.Value)
'    dblOutstandingBalance = objConvAssist.CSafeDbl(adoTotalOutstandingBalance.Value)
'
'    Set adoParameter = Nothing
'    Set adoCommand = Nothing
'
'    'Append to the response
'    Set xmlResponseElem = xmlOut.createElement("MONTHLYREPAYMENT")
'    xmlResponseElem.Text = dblMonthlyRepayment
'    vxmlResponseData.appendChild xmlResponseElem
'
'    'Append to the response
'    Set xmlResponseElem = xmlOut.createElement("OUTSTANDINGBALANCE")
'    xmlResponseElem.Text = dblOutstandingBalance
'    vxmlResponseData.appendChild xmlResponseElem
'
''End:
'
'    Set ICustomerFinancialDO_GetLoansNotRedeemedTotals = vxmlRepsonseData
    
    
    Set xmlCustomerNodeList = Nothing
    Set xmlCustomer = Nothing
    Set adoConnection = Nothing
    Set adoCommand = Nothing
    Set rstThisRecSet = Nothing
    Set adoParameter = Nothing
    Set objIADOAssist = Nothing
    Set objIDOAssist = Nothing
    Set objIomCFClassDef = Nothing
    Set xmlClassDefDoc = Nothing

    If Not m_objContext Is Nothing Then
        m_objContext.SetComplete
    End If
    
    Exit Function

GetLoansNotRedeemedTotalsVbErr:
        
    ' Close the recordset if still open
        If Not rstThisRecSet Is Nothing Then
                If rstThisRecSet.State = adStateOpen Then
                        rstThisRecSet.Close
                End If
        End If

    Set xmlCustomerNodeList = Nothing
    Set xmlCustomer = Nothing
    Set adoConnection = Nothing
    Set adoCommand = Nothing
    Set rstThisRecSet = Nothing
    Set adoParameter = Nothing
    Set objIADOAssist = Nothing
    Set objIDOAssist = Nothing
    Set objIomCFClassDef = Nothing
    Set xmlClassDefDoc = Nothing
    
    m_objErrAssist.AddToErrSource strFunctionName
    
    If m_objErrAssist.IsSystemError = True Then
        m_objErrAssist.LogError TypeName(Me), strFunctionName, Err.Number, Err.Description
    End If
    
    If Not m_objContext Is Nothing Then
        m_objContext.SetComplete
    End If

    '   re-raise error for business object to interpret as appropriate
    Err.Raise Err.Number, Err.Source, Err.Description
End Function

'SR 25/05/2004 : BMIDS772 - new method
Private Function ICustomerFinancialDO_GetCreditHistorySummaryView(ByVal vxmlRequest As IXMLDOMElement) As IXMLDOMNode

On Error GoTo GetCreditHistorySummaryViewVbErr
Const cstrFunctionName As String = "ICustomerFinancialDO_GetCreditHistorySummaryView"

   Dim Cn As ADODB.Connection
    Dim Cmd As ADODB.Command
    Dim objIADOAssist As IADOAssist
     
    Dim xmlOut As FreeThreadedDOMDocument40, xmlResponse As IXMLDOMNode, xmlNode As IXMLDOMNode
    Dim xmlCustNodeList As IXMLDOMNodeList, xmlCustNode As IXMLDOMNode
    Dim xmlFinSummaryNode As IXMLDOMNode
        
    Dim strSql As String, strCustomerNumber As String
    Dim intCustomerVersionNumber As Integer

    Set Cn = New ADODB.Connection
    Set Cmd = New ADODB.Command
    If m_objContext Is Nothing Then
        Set objIADOAssist = New ADOAssist
    Else
        Set objIADOAssist = m_objContext.CreateInstance(gstrBASE_COMPONENT & ".ADOAssist")
    End If
    
    Cn.ConnectionString = objIADOAssist.GetConnStr
    Cn.CursorLocation = adUseClient
    Cn.open
        
    With Cmd
        Set .ActiveConnection = Cn
        .CommandType = adCmdStoredProc
            
        #If GENERIC_SQL Then
            Dim eDbEngineType As DBENGINETYPE
            eDbEngineType = objIADOAssist.GetDBEngine
            If eDbEngineType = SQLServer Then
                strSql = "USP_GETCREDITHISTORYSUMMARYVIEW"
            End If
        #End If
        
        .CommandText = strSql
        .Parameters.Append .CreateParameter(, adBSTR, adParamInput, 12, "")
        .Parameters.Append .CreateParameter(, adInteger, adParamInput, 5, 0)
        .Parameters.Append .CreateParameter(, adInteger, adParamOutput, 1)
        .Parameters.Append .CreateParameter(, adInteger, adParamOutput, 1)
        .Parameters.Append .CreateParameter(, adInteger, adParamOutput, 1)
        .Parameters.Append .CreateParameter(, adInteger, adParamOutput, 1)
        
    End With
        
    Set xmlOut = New FreeThreadedDOMDocument40
    xmlOut.validateOnParse = False
    xmlOut.setProperty "NewParser", True
    Set xmlResponse = xmlOut.createElement("CREDITHISTORYSUMMARYVIEWLIST")
    
    Set xmlCustNodeList = vxmlRequest.selectNodes(".//CUSTOMER")
    For Each xmlCustNode In xmlCustNodeList
        ' Pass appropriate values to INPUT parameters of command object and execute
        strCustomerNumber = m_objXmlAssist.GetTagValue(xmlCustNode, "CUSTOMERNUMBER")
        intCustomerVersionNumber = m_objXmlAssist.GetTagValue(xmlCustNode, "CUSTOMERVERSIONNUMBER")
        
        Cmd(0).Value = strCustomerNumber
        Cmd(1).Value = intCustomerVersionNumber
        Cmd.Execute
        
        Set xmlFinSummaryNode = xmlOut.createElement("CREDITHISTORYSUMMARYVIEW")
        Set xmlNode = xmlOut.createElement("CUSTOMERNUMBER")
        xmlNode.Text = strCustomerNumber
        xmlFinSummaryNode.appendChild xmlNode
        
        Set xmlNode = xmlOut.createElement("CUSTOMERVERSIONNUMBER")
        xmlNode.Text = intCustomerVersionNumber
        xmlFinSummaryNode.appendChild xmlNode
        
        Set xmlNode = xmlOut.createElement("DECLINEDMORTGAGEEXISTINGDATA")
        xmlNode.Text = Cmd(2)
        xmlFinSummaryNode.appendChild xmlNode
        
        Set xmlNode = xmlOut.createElement("ARREARSHISTORYEXISTINGDATA")
        xmlNode.Text = Cmd(3)
        xmlFinSummaryNode.appendChild xmlNode
        
        Set xmlNode = xmlOut.createElement("BANKRUPTCYHISTORYEXISTINGDATA")
        xmlNode.Text = Cmd(4)
        xmlFinSummaryNode.appendChild xmlNode
        
        Set xmlNode = xmlOut.createElement("CCJHISTORYEXISTINGDATA")
        xmlNode.Text = Cmd(5)
        xmlFinSummaryNode.appendChild xmlNode
        
        xmlResponse.appendChild xmlFinSummaryNode
    Next xmlCustNode
        
    Set ICustomerFinancialDO_GetCreditHistorySummaryView = xmlResponse
       
    m_objContext.SetComplete
        
    Exit Function

GetCreditHistorySummaryViewVbErr:

    Set Cn = Nothing
    Set Cmd = Nothing
    Set objIADOAssist = Nothing
    Set xmlOut = Nothing
    Set xmlResponse = Nothing
    Set xmlNode = Nothing
    Set xmlCustNodeList = Nothing
    Set xmlCustNode = Nothing
    Set xmlFinSummaryNode = Nothing
    
    m_objErrAssist.AddToErrSource (cstrFunctionName)
    
    If m_objErrAssist.IsSystemError = True Then
        m_objErrAssist.LogError TypeName(Me), cstrFunctionName, Err.Number, Err.Description
    End If
    
    If m_objErrAssist.IsSystemError Then
        m_objContext.SetAbort
    Else
        m_objContext.SetComplete
    End If
        
    '   re-raise error for business object to interpret as appropriate
    Err.Raise Err.Number, Err.Source, Err.Description

End Function

'SR 25/05/2004 : BMIDS772 - new method
Private Function ICustomerFinancialDO_GetFinancialSummaryView(ByVal vxmlRequest As IXMLDOMElement) As IXMLDOMNode

On Error GoTo GetFinancialSummaryViewVbErr
Const cstrFunctionName As String = "ICustomerFinancialDO_GetFinancialSummaryView"

    Dim Cn As ADODB.Connection
    Dim Cmd As ADODB.Command
    Dim objIADOAssist As IADOAssist
     
    Dim xmlOut As FreeThreadedDOMDocument40, xmlResponse As IXMLDOMNode, xmlNode As IXMLDOMNode
    Dim xmlCustNodeList As IXMLDOMNodeList, xmlCustNode As IXMLDOMNode
    Dim xmlFinSummaryNode As IXMLDOMNode
        
    Dim strSql As String, strCustomerNumber As String
    Dim intCustomerVersionNumber As Integer

    Set Cn = New ADODB.Connection
    Set Cmd = New ADODB.Command
    If m_objContext Is Nothing Then
        Set objIADOAssist = New ADOAssist
    Else
        Set objIADOAssist = m_objContext.CreateInstance(gstrBASE_COMPONENT & ".ADOAssist")
    End If
    
    Cn.ConnectionString = objIADOAssist.GetConnStr
    Cn.CursorLocation = adUseClient
    Cn.open
        
    With Cmd
        Set .ActiveConnection = Cn
        .CommandType = adCmdStoredProc
            
        #If GENERIC_SQL Then
            Dim eDbEngineType As DBENGINETYPE
            eDbEngineType = objIADOAssist.GetDBEngine
            If eDbEngineType = SQLServer Then
                strSql = "USP_GETFINANCIALSUMMARYVIEW"
            End If
        #End If
        
        .CommandText = strSql
        .Parameters.Append .CreateParameter(, adBSTR, adParamInput, 12, "")
        .Parameters.Append .CreateParameter(, adInteger, adParamInput, 5, 0)
        .Parameters.Append .CreateParameter(, adInteger, adParamOutput, 1)
        .Parameters.Append .CreateParameter(, adInteger, adParamOutput, 1)
        .Parameters.Append .CreateParameter(, adInteger, adParamOutput, 1)
        .Parameters.Append .CreateParameter(, adInteger, adParamOutput, 1)
        
    End With
        
    Set xmlOut = New FreeThreadedDOMDocument40
    xmlOut.validateOnParse = False
    xmlOut.setProperty "NewParser", True
    Set xmlResponse = xmlOut.createElement("FINANCIALSUMMARYVIEWLIST")
    
    Set xmlCustNodeList = vxmlRequest.selectNodes(".//CUSTOMER")
    For Each xmlCustNode In xmlCustNodeList
        ' Pass appropriate values to INPUT parameters of command object and execute
        strCustomerNumber = m_objXmlAssist.GetTagValue(xmlCustNode, "CUSTOMERNUMBER")
        intCustomerVersionNumber = m_objXmlAssist.GetTagValue(xmlCustNode, "CUSTOMERVERSIONNUMBER")
        
        Cmd(0).Value = strCustomerNumber
        Cmd(1).Value = intCustomerVersionNumber
        Cmd.Execute
        
        Set xmlFinSummaryNode = xmlOut.createElement("FINANCIALSUMMARYVIEW")
        Set xmlNode = xmlOut.createElement("CUSTOMERNUMBER")
        xmlNode.Text = strCustomerNumber
        xmlFinSummaryNode.appendChild xmlNode
        
        Set xmlNode = xmlOut.createElement("CUSTOMERVERSIONNUMBER")
        xmlNode.Text = intCustomerVersionNumber
        xmlFinSummaryNode.appendChild xmlNode
        
        Set xmlNode = xmlOut.createElement("REGULAROUTGOINGSEXISTINGDATA")
        xmlNode.Text = Cmd(2)
        xmlFinSummaryNode.appendChild xmlNode
        
        Set xmlNode = xmlOut.createElement("MORTGAGEACCOUNTEXISTINGDATA")
        xmlNode.Text = Cmd(3)
        xmlFinSummaryNode.appendChild xmlNode
        
        Set xmlNode = xmlOut.createElement("LOANSLIABILITIESEXISTINGDATA")
        xmlNode.Text = Cmd(4)
        xmlFinSummaryNode.appendChild xmlNode
        
        Set xmlNode = xmlOut.createElement("BANKCREDITCARDEXISTINGDATA")
        xmlNode.Text = Cmd(5)
        xmlFinSummaryNode.appendChild xmlNode
        
        xmlResponse.appendChild xmlFinSummaryNode
    Next xmlCustNode
    
    Set ICustomerFinancialDO_GetFinancialSummaryView = xmlResponse
       
    m_objContext.SetComplete
        
    Exit Function

GetFinancialSummaryViewVbErr:

    Set Cn = Nothing
    Set Cmd = Nothing
    Set objIADOAssist = Nothing
    Set xmlOut = Nothing
    Set xmlResponse = Nothing
    Set xmlNode = Nothing
    Set xmlCustNodeList = Nothing
    Set xmlCustNode = Nothing
    Set xmlFinSummaryNode = Nothing

    m_objErrAssist.AddToErrSource (cstrFunctionName)
    
    If m_objErrAssist.IsSystemError = True Then
        m_objErrAssist.LogError TypeName(Me), cstrFunctionName, Err.Number, Err.Description
    End If
    
    If m_objErrAssist.IsSystemError Then
        m_objContext.SetAbort
    Else
        m_objContext.SetComplete
    End If
        
    '   re-raise error for business object to interpret as appropriate
    Err.Raise Err.Number, Err.Source, Err.Description

End Function

'BMIDS00336 MDC 10/09/2002
Private Function ICustomerFinancialDO_GetMortgageAccountLoanTotals(ByVal vxmlTableElement As IXMLDOMElement) As IXMLDOMNode
' header ----------------------------------------------------------------------------------
' procedure:    ICustomerFinancialDO_GetMortgageAccountLoanTotals
' author:       Mark Coleman
' date:         10/09/2002
' description:  Retrieves Mortgage Loan totals for a given Mortgage Account
' pass:         vxmlTableElement  XML Input request specifying the data and operation
' return:       XML Response containing single record for given Mortgage Account
'------------------------------------------------------------------------------------------

On Error GoTo GetMortgageAccountLoanTotalsVbErr

    
Dim objErrAssist As New ErrAssist
Dim objIDOAssist As IDOAssist

Dim xmlClassDefDoc As FreeThreadedDOMDocument40
Dim objClassDef As ICFClassDef
Dim xmlReturn As IXMLDOMNode

Const strFunctionName As String = "ICustomerFinancialDO_GetMortgageAccountLoanTotals"
Dim objContext As ObjectContext
    
    Set objContext = GetObjectContext()
            
    If objContext Is Nothing Then
        Set objIDOAssist = New DOAssist
    Else
        Set objIDOAssist = objContext.CreateInstance(gstrBASE_COMPONENT & ".DOAssist")
    End If
    
    Set objClassDef = New CFClassDef
    Set xmlClassDefDoc = objClassDef.LoadMortgageAccountLoanTotalsData()
    Set xmlReturn = objIDOAssist.GetData(vxmlTableElement, xmlClassDefDoc)
    
    Set ICustomerFinancialDO_GetMortgageAccountLoanTotals = xmlReturn

    Set objClassDef = Nothing
    Set xmlClassDefDoc = Nothing
    Set objIDOAssist = Nothing
    Set objErrAssist = Nothing
    Set xmlReturn = Nothing
    
    If Not objContext Is Nothing Then
        objContext.SetComplete
    End If
    
    Exit Function

GetMortgageAccountLoanTotalsVbErr:
    
    Set objClassDef = Nothing
    Set xmlClassDefDoc = Nothing
    Set objIDOAssist = Nothing
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
'BMIDS00336 MDC 10/09/2002 - End

Private Function ICustomerFinancialDO_GetMortgageLenderDirectoryData(ByVal vxmlTableElement As IXMLDOMElement) As IXMLDOMNode
' header ----------------------------------------------------------------------------------
' description:
'   Get the data for a single instance of the persistant data associated with
'   this data object
' pass:
'   vxmlxmlTableElement  xml element containing the request
' return:                xml node containing retrieved data
'------------------------------------------------------------------------------------------
On Error GoTo GetMortgageLenderDirectoryDataVbErr

    Const strFunctionName As String = "ICustomerFinancialDO_GetMortgageLenderDirectoryData"
    
    Dim objErrAssist As New ErrAssist
    
    Dim objIDOAssist As IDOAssist, _
        objTPAssist As IThirdPartyAssist
    Set objTPAssist = New ThirdPartyAssist

    Dim xmlClassDefDoc As FreeThreadedDOMDocument40
    Dim objClassDef As ICFClassDef
    Dim xmlReturn As IXMLDOMNode
    
    Dim objContext As ObjectContext
    Set objContext = GetObjectContext()
            
    If objContext Is Nothing Then
        Set objIDOAssist = New DOAssist
    Else
        Set objIDOAssist = objContext.CreateInstance(gstrBASE_COMPONENT & ".DOAssist")
    End If
    
    Set objClassDef = New CFClassDef
    Set xmlClassDefDoc = objClassDef.LoadMortgageLenderDirectoryData
    Set xmlReturn = objIDOAssist.GetData(vxmlTableElement, xmlClassDefDoc)
    
   ' objTPAssist.GetThirdParty xmlReturn

    Set ICustomerFinancialDO_GetMortgageLenderDirectoryData = xmlReturn

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

GetMortgageLenderDirectoryDataVbErr:
    
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

Private Function ICustomerFinancialDO_GetMortgageAccount(ByVal vxmlTableElement As IXMLDOMElement) As IXMLDOMNode

' header ----------------------------------------------------------------------------------
' description:
'   Get the data for a single instance of the persistant data associated with
'   this data object
' pass:
'   vxmlxmlTableElement  xml element containing the request
' return:                xml node containing retrieved data
'------------------------------------------------------------------------------------------
On Error GoTo GetMortgageAccountVbErr

    Const strFunctionName As String = "ICustomerFinancialDO_GetMortgageAccount"
    
    Dim objErrAssist As New ErrAssist
    
    Dim objIDOAssist As IDOAssist, _
        objXmlAssist As New XMLAssist
    Dim xmlClassDefDoc As FreeThreadedDOMDocument40
    Dim objICFClassDef As ICFClassDef
    
    Dim xmlReturn As IXMLDOMNode
    
    Dim objContext As ObjectContext
    Set objContext = GetObjectContext()
            
    If objContext Is Nothing Then
        Set objIDOAssist = New DOAssist
    Else
        Set objIDOAssist = objContext.CreateInstance(gstrBASE_COMPONENT & ".DOAssist")
    End If
    
    Set objICFClassDef = New CFClassDef
    Set xmlClassDefDoc = objICFClassDef.LoadMortgageAccDetsData
    'Set xmlReturn = objIDOAssist.GetData(vxmlTableElement, xmlClassDefDoc, vstrItemName:="THIRDPARTY")
    Set xmlReturn = objIDOAssist.GetData(vxmlTableElement, xmlClassDefDoc)
    '''GetData = objDOAssist.GetData(vstrXMLRequest, objICMClassDef.LoadPPSubQuoteData)
    
    'objXmlAssist.PromoteNode xmlReturn, xmlReturn.selectSingleNode("THIRDPARTY")

    Set ICustomerFinancialDO_GetMortgageAccount = xmlReturn

    Set objICFClassDef = Nothing
    Set xmlClassDefDoc = Nothing
    Set objIDOAssist = Nothing
    Set objErrAssist = Nothing
    Set xmlReturn = Nothing
    Set objXmlAssist = Nothing
    
    If Not objContext Is Nothing Then
        objContext.SetComplete
    End If
    
    Exit Function

GetMortgageAccountVbErr:
    
    Set objICFClassDef = Nothing
    Set xmlClassDefDoc = Nothing
    Set objIDOAssist = Nothing
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

Private Sub ICustomerFinancialDO_CreateArrearsHistory(ByVal vxmlTableElement As IXMLDOMElement)
' header ----------------------------------------------------------------------------------
' description:
'   CreateArrearsHistory an instance of the persistant data associated with this data object
'   for each set of data in the request
' pass:
'   vxmlxmlTableElement  xml element containing the request
' return:                n/a
'------------------------------------------------------------------------------------------
On Error GoTo CreateArrearsHistoryVbErr

    Const strFunctionName As String = "ICustomerFinancialDO_CreateArrearsHistory"
   
    Dim xmlClassDefDoc          As FreeThreadedDOMDocument40
    Dim objICFClassDef          As ICFClassDef
    'BMISD00190
    Dim xmlAccountRelClassDef   As FreeThreadedDOMDocument40
    Dim xmlOtherAAClassDef      As FreeThreadedDOMDocument40
    Dim xmlAccountClassDef      As FreeThreadedDOMDocument40
    Dim xmlRequest              As IXMLDOMNode
    Dim xmlCustomer             As IXMLDOMNode
    Dim xmlCopy                 As IXMLDOMNode
    Dim xmlGUIDNode             As IXMLDOMNode
    Dim xmlAccountGuid          As IXMLDOMNode
    Dim xmlOtherArrears         As IXMLDOMNode
    'BMIDS00190 End
    
    Set objICFClassDef = New CFClassDef
    Set xmlClassDefDoc = objICFClassDef.LoadArrearsHistoryData()

    'BMIDS00190

    Set xmlOtherArrears = vxmlTableElement.selectSingleNode("OTHERARREARSACCOUNT")
    
    'Save customers linked to the Arrears History for Other Arrears Account
    'This is not necessary to MortgageAccount or LoansLiabilities as they are already linked
    If Not xmlOtherArrears Is Nothing Then
        
        Set xmlAccountClassDef = objICFClassDef.LoadAccountData()
        Set xmlAccountRelClassDef = objICFClassDef.LoadAccountRelationshipData()
        Set xmlOtherAAClassDef = objICFClassDef.LoadOtherArrearsAccountData()
        
        'Create a new Account
        Set xmlRequest = xmlClassDefDoc.createElement("ACCOUNT")
        Set xmlGUIDNode = xmlClassDefDoc.createElement("ACCOUNTGUID")
        xmlGUIDNode.Text = CreateGUID()
        xmlRequest.appendChild xmlGUIDNode.cloneNode(True)
        
        m_objIDOAssist.Create xmlRequest, xmlAccountClassDef
        
        'Create a new OtherArrearsAccount
        xmlOtherArrears.appendChild xmlGUIDNode.cloneNode(True)
        m_objIDOAssist.Create xmlOtherArrears, xmlOtherAAClassDef
        vxmlTableElement.removeChild xmlOtherArrears
        
        'Set the new AccountGUID node correctly for ArrearsHistory
        Set xmlAccountGuid = vxmlTableElement.selectSingleNode("ACCOUNTGUID")
        If xmlAccountGuid Is Nothing Then
            vxmlTableElement.appendChild xmlGUIDNode.cloneNode(True)
        Else
            xmlAccountGuid.Text = xmlGUIDNode.Text
        End If
        
        'Create new AccountRelationships per customer
        For Each xmlCustomer In vxmlTableElement.selectNodes("ACCOUNTRELATIONSHIP")
            xmlCustomer.appendChild xmlGUIDNode.cloneNode(True)
            m_objIDOAssist.Create xmlCustomer, xmlAccountRelClassDef
            vxmlTableElement.removeChild xmlCustomer
        Next
    
    End If
        
    Set xmlGUIDNode = xmlClassDefDoc.createElement("ARREARSHISTORYGUID")
    xmlGUIDNode.Text = CreateGUID()
    vxmlTableElement.appendChild xmlGUIDNode

    'm_objIDOAssist.GenerateSequenceNumber vxmlTableElement, xmlClassDefDoc, "ARREARSSEQUENCENUMBER"
    m_objIDOAssist.Create vxmlTableElement, xmlClassDefDoc
    
    Set xmlAccountRelClassDef = Nothing
    Set xmlOtherAAClassDef = Nothing
    Set xmlAccountClassDef = Nothing
    Set xmlRequest = Nothing
    Set xmlCustomer = Nothing
    Set xmlCopy = Nothing
    Set xmlGUIDNode = Nothing
    Set xmlAccountGuid = Nothing
    Set xmlOtherArrears = Nothing
    'BMIDS00190 End
    
    Set objICFClassDef = Nothing
    Set xmlClassDefDoc = Nothing

    If Not m_objContext Is Nothing Then
        m_objContext.SetComplete
    End If

    Exit Sub

CreateArrearsHistoryVbErr:
    
    'BMIDS00190
    Set xmlAccountRelClassDef = Nothing
    Set xmlOtherAAClassDef = Nothing
    Set xmlAccountClassDef = Nothing
    Set xmlRequest = Nothing
    Set xmlCustomer = Nothing
    Set xmlCopy = Nothing
    Set xmlGUIDNode = Nothing
    Set xmlAccountGuid = Nothing
    Set xmlOtherArrears = Nothing
    'BMIDS00190 End
    Set objICFClassDef = Nothing
    Set xmlClassDefDoc = Nothing
    
    m_objErrAssist.AddToErrSource strFunctionName
    
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

Private Sub ICustomerFinancialDO_CreateBankCard(ByVal vxmlTableElement As IXMLDOMElement)
' header ----------------------------------------------------------------------------------
' description:
'   CreateBankCard an instance of the persistant data associated with this data object
'   for each set of data in the request
' pass:
'   vxmlxmlTableElement  xml element containing the request
' return:                n/a
'------------------------------------------------------------------------------------------
On Error GoTo CreateBankCardVbErr

    Const strFunctionName As String = "ICustomerFinancialDO_CreateBankCard"
   
    Dim objErrAssist As New ErrAssist
    Dim xmlClassDefDoc As FreeThreadedDOMDocument40
    Dim objICFClassDef As ICFClassDef
    
    Dim objIDOAssist As IDOAssist
    
    Dim objContext As ObjectContext
    Set objContext = GetObjectContext()
                    
    If objContext Is Nothing Then
        Set objIDOAssist = New DOAssist
    Else
        Set objIDOAssist = objContext.CreateInstance(gstrBASE_COMPONENT & ".DOAssist")
    End If
    
    Set objICFClassDef = New CFClassDef
    Set xmlClassDefDoc = objICFClassDef.LoadBankCreditCardData

    objIDOAssist.GenerateSequenceNumber vxmlTableElement, xmlClassDefDoc, "SEQUENCENUMBER"
    objIDOAssist.Create vxmlTableElement, xmlClassDefDoc
    
    Set objICFClassDef = Nothing
    Set xmlClassDefDoc = Nothing
    Set objIDOAssist = Nothing
    Set objErrAssist = Nothing

    If Not objContext Is Nothing Then
        objContext.SetComplete
    End If

    Exit Sub

CreateBankCardVbErr:
    
    Set objICFClassDef = Nothing
    Set xmlClassDefDoc = Nothing
    Set objIDOAssist = Nothing
    
    objErrAssist.AddToErrSource strFunctionName
    
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

Private Sub ICustomerFinancialDO_CreateBankruptcyHistory(ByVal vxmlTableElement As IXMLDOMElement)
' header ----------------------------------------------------------------------------------
' description:
'   CreateBankruptcyHistory an instance of the persistant data associated with this data object
'   for each set of data in the request
' pass:
'   vxmlxmlTableElement  xml element containing the request
' return:                n/a
'------------------------------------------------------------------------------------------
On Error GoTo CreateBankruptcyHistoryVbErr

    Const strFunctionName   As String = "ICustomerFinancialDO_CreateBankruptcyHistory"
   
    Dim xmlClassDefDoc      As FreeThreadedDOMDocument40
    Dim objICFClassDef      As ICFClassDef
    'BMISD00190
    Dim xmlCVBHClassDefDoc  As FreeThreadedDOMDocument40
    Dim xmlRequest          As IXMLDOMElement
    Dim xmlCustomer         As IXMLDOMNode
    Dim xmlCopy             As IXMLDOMElement
    Dim xmlGUIDNode         As IXMLDOMElement
    'BMIDS00190 End
    
    Set objICFClassDef = New CFClassDef
    Set xmlClassDefDoc = objICFClassDef.LoadBankruptcyHistoryData()
    
    'BMIDS00190
    'Create a copy of vxmlTableElement and then remove the Customer Nodes
    Set xmlCopy = vxmlTableElement.cloneNode(True)
    
    For Each xmlCustomer In xmlCopy.selectNodes("CUSTOMERVERSIONBANKRUPTCYHISTORY")
        xmlCopy.removeChild xmlCustomer
    Next
    
    Set xmlGUIDNode = xmlClassDefDoc.createElement("BANKRUPTCYHISTORYGUID")
    xmlGUIDNode.Text = CreateGUID()
    
    xmlCopy.appendChild xmlGUIDNode.cloneNode(True)
    
    'm_objIDOAssist.GenerateSequenceNumber vxmlTableElement, xmlClassDefDoc, "SEQUENCENUMBER"
    m_objIDOAssist.Create xmlCopy, xmlClassDefDoc
    
    Set xmlCVBHClassDefDoc = objICFClassDef.LoadCustVerBankruptcyHistoryData()
    For Each xmlCustomer In vxmlTableElement.selectNodes("CUSTOMERVERSIONBANKRUPTCYHISTORY")
        xmlCustomer.appendChild xmlGUIDNode.cloneNode(True)
        m_objIDOAssist.Create xmlCustomer, xmlCVBHClassDefDoc
    Next
    
    Set xmlCVBHClassDefDoc = Nothing
    Set xmlRequest = Nothing
    Set xmlCustomer = Nothing
    Set xmlCopy = Nothing
    Set xmlGUIDNode = Nothing
    'BMIDS00190 End
    
    Set objICFClassDef = Nothing
    Set xmlClassDefDoc = Nothing

    If Not m_objContext Is Nothing Then
        m_objContext.SetComplete
    End If

    Exit Sub

CreateBankruptcyHistoryVbErr:
    
    'BMIDS00190
    Set xmlCVBHClassDefDoc = Nothing
    Set xmlRequest = Nothing
    Set xmlCustomer = Nothing
    Set xmlCopy = Nothing
    Set xmlGUIDNode = Nothing
    'BMIDS00190 End
    
    Set objICFClassDef = Nothing
    Set xmlClassDefDoc = Nothing
    
    m_objErrAssist.AddToErrSource strFunctionName
    
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

Private Sub ICustomerFinancialDO_CreateCCJHistory(ByVal vxmlTableElement As IXMLDOMElement)
' header ----------------------------------------------------------------------------------
' description:
'   CreateCCJHistory an instance of the persistant data associated with this data object
'   for each set of data in the request
' pass:
'   vxmlxmlTableElement  xml element containing the request
' return:                n/a
'------------------------------------------------------------------------------------------
On Error GoTo CreateCCJHistoryVbErr

    Const strFunctionName   As String = "ICustomerFinancialDO_CreateCCJHistory"
   
    Dim xmlClassDefDoc      As FreeThreadedDOMDocument40
    Dim objICFClassDef      As ICFClassDef
    'BMISD00190
    Dim xmlCVCHClassDefDoc  As FreeThreadedDOMDocument40
    Dim xmlRequest          As IXMLDOMElement
    Dim xmlCustomer         As IXMLDOMNode
    Dim xmlCopy             As IXMLDOMElement
    Dim xmlGUIDNode         As IXMLDOMElement
    'BMIDS00190 End
    
    Set objICFClassDef = New CFClassDef
    Set xmlClassDefDoc = objICFClassDef.LoadCCJHistoryData()

    'BMIDS00190
    'Create a copy of vxmlTableElement and then remove the Customer Nodes
    Set xmlCopy = vxmlTableElement.cloneNode(True)
    
    For Each xmlCustomer In xmlCopy.selectNodes("CUSTOMERVERSIONCCJHISTORY")
        xmlCopy.removeChild xmlCustomer
    Next
    
    Set xmlGUIDNode = xmlClassDefDoc.createElement("CCJHISTORYGUID")
    xmlGUIDNode.Text = CreateGUID()
    
    xmlCopy.appendChild xmlGUIDNode.cloneNode(True)

    'm_objIDOAssist.GenerateSequenceNumber vxmlTableElement, xmlClassDefDoc, "SEQUENCENUMBER"
    m_objIDOAssist.Create xmlCopy, xmlClassDefDoc
    
    Set xmlCVCHClassDefDoc = objICFClassDef.LoadCustVerCCJHistoryData()
    For Each xmlCustomer In vxmlTableElement.selectNodes("CUSTOMERVERSIONCCJHISTORY")
        xmlCustomer.appendChild xmlGUIDNode.cloneNode(True)
        m_objIDOAssist.Create xmlCustomer, xmlCVCHClassDefDoc
    Next
    
    Set xmlCVCHClassDefDoc = Nothing
    Set xmlRequest = Nothing
    Set xmlCustomer = Nothing
    Set xmlCopy = Nothing
    Set xmlGUIDNode = Nothing
    'BMIDS00190 End
    
    Set objICFClassDef = Nothing
    Set xmlClassDefDoc = Nothing

    If Not m_objContext Is Nothing Then
        m_objContext.SetComplete
    End If

    Exit Sub

CreateCCJHistoryVbErr:
    
    'BMIDS00190
    Set xmlCVCHClassDefDoc = Nothing
    Set xmlRequest = Nothing
    Set xmlCustomer = Nothing
    Set xmlCopy = Nothing
    Set xmlGUIDNode = Nothing
    'BMIDS00190 End
    
    Set objICFClassDef = Nothing
    Set xmlClassDefDoc = Nothing
    
    m_objErrAssist.AddToErrSource strFunctionName
    
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

Private Sub ICustomerFinancialDO_CreateDeclinedMortgage(ByVal vxmlTableElement As IXMLDOMElement)
' header ----------------------------------------------------------------------------------
' description:
'   CreateDeclinedMortgage an instance of the persistant data associated with this data object
'   for each set of data in the request
' pass:
'   vxmlxmlTableElement  xml element containing the request
' return:                n/a
'------------------------------------------------------------------------------------------
On Error GoTo CreateDeclinedMortgageVbErr

    Const strFunctionName As String = "ICustomerFinancialDO_CreateDeclinedMortgage"
   
    Dim objErrAssist As New ErrAssist
    Dim xmlClassDefDoc As FreeThreadedDOMDocument40
    Dim objICFClassDef As ICFClassDef
    
    Dim objIDOAssist As IDOAssist
    
    Dim objContext As ObjectContext
    Set objContext = GetObjectContext()
                    
    If objContext Is Nothing Then
        Set objIDOAssist = New DOAssist
    Else
        Set objIDOAssist = objContext.CreateInstance(gstrBASE_COMPONENT & ".DOAssist")
    End If
    
    Set objICFClassDef = New CFClassDef
    Set xmlClassDefDoc = objICFClassDef.LoadDeclinedMortgageData()

    objIDOAssist.GenerateSequenceNumber vxmlTableElement, xmlClassDefDoc, "SEQUENCENUMBER"
    objIDOAssist.Create vxmlTableElement, xmlClassDefDoc
    
    Set objICFClassDef = Nothing
    Set xmlClassDefDoc = Nothing
    Set objIDOAssist = Nothing
    Set objErrAssist = Nothing

    If Not objContext Is Nothing Then
        objContext.SetComplete
    End If

    Exit Sub

CreateDeclinedMortgageVbErr:
    
    Set objICFClassDef = Nothing
    Set xmlClassDefDoc = Nothing
    Set objIDOAssist = Nothing
    
    objErrAssist.AddToErrSource strFunctionName
    
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

Private Sub ICustomerFinancialDO_CreateGroupConnection(ByVal vxmlTableElement As IXMLDOMElement)
' header ----------------------------------------------------------------------------------
' description:
'   CreateGroupConnection an instance of the persistant data associated with this data object
'   for each set of data in the request
' pass:
'   vxmlxmlTableElement  xml element containing the request
' return:                n/a
'------------------------------------------------------------------------------------------
On Error GoTo CreateGroupConnectionVbErr

    Const strFunctionName As String = "ICustomerFinancialDO_CreateGroupConnection"
   
    Dim objErrAssist As New ErrAssist
    Dim xmlClassDefDoc As FreeThreadedDOMDocument40
    Dim objICFClassDef As ICFClassDef
    
    Dim objIDOAssist As IDOAssist
    
    Dim objContext As ObjectContext
    Set objContext = GetObjectContext()
                    
    If objContext Is Nothing Then
        Set objIDOAssist = New DOAssist
    Else
        Set objIDOAssist = objContext.CreateInstance(gstrBASE_COMPONENT & ".DOAssist")
    End If
    
    Set objICFClassDef = New CFClassDef
    Set xmlClassDefDoc = objICFClassDef.LoadGroupConnectionData()

    objIDOAssist.GenerateSequenceNumber vxmlTableElement, xmlClassDefDoc, "GROUPCONNECTIONSEQUENCENUMBER"
    objIDOAssist.Create vxmlTableElement, xmlClassDefDoc
    
    Set objICFClassDef = Nothing
    Set xmlClassDefDoc = Nothing
    Set objIDOAssist = Nothing
    Set objErrAssist = Nothing

    If Not objContext Is Nothing Then
        objContext.SetComplete
    End If

    Exit Sub

CreateGroupConnectionVbErr:
    
    Set objICFClassDef = Nothing
    Set xmlClassDefDoc = Nothing
    Set objIDOAssist = Nothing
    
    objErrAssist.AddToErrSource strFunctionName
    
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

Private Sub ICustomerFinancialDO_CreateLiability(ByVal vxmlTableElement As IXMLDOMElement)
' header ----------------------------------------------------------------------------------
' description:
'   CreateLiability an instance of the persistant data associated with this data object
'   for each set of data in the request
' pass:
'   vxmlxmlTableElement  xml element containing the request
' return:                n/a
'------------------------------------------------------------------------------------------
On Error GoTo CreateLiabilityVbErr

    Const strFunctionName As String = "ICustomerFinancialDO_CreateLiability"
   
    Dim objTPAssist   As IThirdPartyAssist
    Set objTPAssist = New ThirdPartyAssist
    
    Dim xmlClassDefDoc As FreeThreadedDOMDocument40
    Dim objICFClassDef As ICFClassDef
    
    'BMIDS00190
    Dim xmlAccountClassDefDoc       As FreeThreadedDOMDocument40
    Dim xmlAccountRelClassDefDoc    As FreeThreadedDOMDocument40
    Dim xmlGUIDNode                 As IXMLDOMNode
    Dim xmlCustomer                 As IXMLDOMNode
    Dim xmlAccount                  As IXMLDOMNode
    Dim xmlThirdPartyGUID           As IXMLDOMNode
    Dim xmlDirectoryGUID            As IXMLDOMNode
    'BMIDS00190 End

    Dim objThirdPartyBO As ThirdPartyBO
    Dim strRequest As String
    
    If m_objContext Is Nothing Then
        Set objThirdPartyBO = New ThirdPartyBO
    Else
        Set objThirdPartyBO = m_objContext.CreateInstance(gstrTHIRDPARTY_COMPONENT & ".ThirdPartyBO")
    End If
    
    Set objICFClassDef = New CFClassDef
    Set xmlClassDefDoc = objICFClassDef.LoadLoansLiabilitiesData()

    'BMIDS00190
    Set xmlAccount = vxmlTableElement.selectSingleNode("ACCOUNT")
    Set xmlThirdPartyGUID = xmlAccount.selectSingleNode("THIRDPARTYGUID")
    Set xmlDirectoryGUID = xmlAccount.selectSingleNode("DIRECTORYGUID")
    
    'Move ThirdPartyGUID and DirectoryGUID one level up for MaintainThirdParty
    vxmlTableElement.appendChild xmlThirdPartyGUID
    vxmlTableElement.appendChild xmlDirectoryGUID
    'BMIDS00190 End

    ' Maintain THIRDPARTY if necessary
    If (Trim$(m_objXmlAssist.GetTagValue(vxmlTableElement, "THIRDPARTY")) <> "") And _
       (Trim$(m_objXmlAssist.GetTagValue(vxmlTableElement, "DIRECTORYGUID")) = "") Then
        ' THIRDPARTY details exist in the XML, so maintain the appropriate THIRDPARTY record
        objTPAssist.MaintainThirdParty vxmlTableElement
    End If

    'BMIDS00190
    Set xmlAccountClassDefDoc = objICFClassDef.LoadAccountData()
    Set xmlAccountRelClassDefDoc = objICFClassDef.LoadAccountRelationshipData()
    
    Set xmlGUIDNode = xmlClassDefDoc.createElement("ACCOUNTGUID")
    xmlGUIDNode.Text = CreateGUID()
    
    'Move ThirdPartyGUID and DirectoryGUID back to Account node
    xmlAccount.appendChild xmlThirdPartyGUID
    xmlAccount.appendChild xmlDirectoryGUID
    
    'Create Account
    xmlAccount.appendChild xmlGUIDNode.cloneNode(True)
    m_objIDOAssist.Create xmlAccount, xmlAccountClassDefDoc
    
    'Create AccountRelationships per customer
    For Each xmlCustomer In vxmlTableElement.selectNodes("ACCOUNTRELATIONSHIP")
        xmlCustomer.appendChild xmlGUIDNode.cloneNode(True)
        m_objIDOAssist.Create xmlCustomer, xmlAccountRelClassDefDoc
        vxmlTableElement.removeChild xmlCustomer
    Next
    
    vxmlTableElement.appendChild xmlGUIDNode
    'BMIDS00190 End
    
    '
    ' Create LOANSLIABILITIES record
    '
    'BMIDS00190
    'm_objIDOAssist.GenerateSequenceNumber vxmlTableElement, xmlClassDefDoc, "SEQUENCENUMBER"
    vxmlTableElement.removeChild xmlAccount
    'BMIDS00190 End
    m_objIDOAssist.Create vxmlTableElement, xmlClassDefDoc
    
    Set objICFClassDef = Nothing
    Set xmlClassDefDoc = Nothing
    Set objThirdPartyBO = Nothing
    Set objTPAssist = Nothing

    If Not m_objContext Is Nothing Then
        m_objContext.SetComplete
    End If

    Exit Sub

CreateLiabilityVbErr:
    
    Set objICFClassDef = Nothing
    Set xmlClassDefDoc = Nothing
    Set objThirdPartyBO = Nothing
    Set objTPAssist = Nothing
    
    m_objErrAssist.AddToErrSource strFunctionName
    
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

Private Sub ICustomerFinancialDO_CreateLifeProduct(ByVal vxmlTableElement As IXMLDOMElement, _
                                                   Optional ByVal vblnSimpleCreate As Boolean = False)
' header ----------------------------------------------------------------------------------
' description:
'   CreateLifeProduct an instance of the persistant data associated with this data object
'   for each set of data in the request
' pass:
'   vxmlxmlTableElement  xml element containing the request
' return:                n/a
'------------------------------------------------------------------------------------------
On Error GoTo CreateLifeProductVbErr

    Const strFunctionName As String = "ICustomerFinancialDO_CreateLifeProduct"
   
    Dim objErrAssist  As ErrAssist, _
        objXmlAssist  As XMLAssist, _
        objTPAssist   As IThirdPartyAssist
    Set objErrAssist = New ErrAssist
    Set objXmlAssist = New XMLAssist
    Set objTPAssist = New ThirdPartyAssist
    Dim xmlClassDefDoc As FreeThreadedDOMDocument40
    Dim objICFClassDef As ICFClassDef

    Dim xmlAccountNode            As IXMLDOMNode, _
        xmlLifePolicyNode         As IXMLDOMNode, _
        xmlACNode         As IXMLDOMNode, _
        xmlMRCNode                As IXMLDOMNode

    Dim strAccountGUID As String
    
    Dim objIDOAssist As IDOAssist
    
    Dim objContext As ObjectContext
    Set objContext = GetObjectContext()
                    
    If objContext Is Nothing Then
        Set objIDOAssist = New DOAssist
    Else
        Set objIDOAssist = objContext.CreateInstance(gstrBASE_COMPONENT & ".DOAssist")
    End If
    
    Set objICFClassDef = New CFClassDef
    strAccountGUID = CreateGUID()

    Set xmlAccountNode = vxmlTableElement.selectSingleNode(".//ACCOUNT")
    Set xmlLifePolicyNode = vxmlTableElement.selectSingleNode(".//LIFEPOLICY")
    Set xmlMRCNode = vxmlTableElement.selectSingleNode(".//MORTGAGERELATEDCONTRACTS")
    Set xmlACNode = vxmlTableElement.selectSingleNode(".//APPLICATIONCONTRACT")

    If Not vblnSimpleCreate Then
        '
        ' Create ACCOUNT record
        '
        If Not (xmlAccountNode Is Nothing) Then
            xmlAccountNode.appendChild(xmlAccountNode.ownerDocument.createElement("ACCOUNTGUID")).Text = strAccountGUID
        
            ' Maintain THIRDPARTY if necessary
            If (Trim$(objXmlAssist.GetTagValue(vxmlTableElement, "THIRDPARTY")) <> "") And _
               (Trim$(objXmlAssist.GetTagValue(vxmlTableElement, "DIRECTORYGUID")) = "") Then
                ' THIRDPARTY details exist in the XML, so maintain the appropriate THIRDPARTY record
                objTPAssist.MaintainThirdParty xmlAccountNode
            End If
            objIDOAssist.Create xmlAccountNode, objICFClassDef.LoadAccountData
        End If

        '
        ' Create LIFEPOLICY record
        '
        If Not (xmlLifePolicyNode Is Nothing) Then
            xmlLifePolicyNode.appendChild(xmlLifePolicyNode.ownerDocument.createElement("ACCOUNTGUID")).Text = strAccountGUID
            objIDOAssist.Create xmlLifePolicyNode, objICFClassDef.LoadLifePolicyData
        End If
        
        '
        ' Create MORTGAGERELATEDCONTRACTS record
        '
        If Not (xmlMRCNode Is Nothing) Then
            xmlMRCNode.appendChild(xmlMRCNode.ownerDocument.createElement("ACCOUNTGUID")).Text = strAccountGUID
            objIDOAssist.Create xmlMRCNode, objICFClassDef.LoadMortgageRelatedContractsData
        End If
        
        '
        ' Create ApplicationContract record
        '
        If Not (xmlACNode Is Nothing) Then
            xmlACNode.appendChild(xmlACNode.ownerDocument.createElement("ACCOUNTGUID")).Text = strAccountGUID
            objIDOAssist.Create xmlACNode, objICFClassDef.LoadApplicationContractData
        End If
    End If


    Set objICFClassDef = Nothing
    Set xmlClassDefDoc = Nothing
    Set objIDOAssist = Nothing
    Set objErrAssist = Nothing
    Set xmlAccountNode = Nothing
    Set xmlLifePolicyNode = Nothing
    Set xmlMRCNode = Nothing
    Set objXmlAssist = Nothing
    Set objTPAssist = Nothing

    If Not objContext Is Nothing Then
        objContext.SetComplete
    End If

    Exit Sub

CreateLifeProductVbErr:
    
    Set objICFClassDef = Nothing
    Set xmlClassDefDoc = Nothing
    Set objIDOAssist = Nothing
    Set xmlAccountNode = Nothing
    Set xmlLifePolicyNode = Nothing
    Set xmlMRCNode = Nothing
    Set objXmlAssist = Nothing
    Set objTPAssist = Nothing
    
    objErrAssist.AddToErrSource strFunctionName
    
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

Private Sub ICustomerFinancialDO_CreateMortgageAccount(ByVal vxmlTableElement As IXMLDOMElement, _
                                                       ByRef rxmlKeysNode As IXMLDOMNode)
' header ----------------------------------------------------------------------------------
' description:
'   CreateMortgageAccount an instance of the persistant data associated with this data object
'   for each set of data in the request
' pass:
'   vxmlxmlTableElement  xml element containing the request
' return:                n/a
'------------------------------------------------------------------------------------------
On Error GoTo CreateMortgageAccountVbErr

    Const strFunctionName As String = "ICustomerFinancialDO_CreateMortgageAccount"
   
    Dim objErrAssist  As New ErrAssist, _
        objXmlAssist  As XMLAssist, _
        objTPAssist   As IThirdPartyAssist
    Set objTPAssist = New ThirdPartyAssist
    Set objXmlAssist = New XMLAssist
    Dim xmlClassDefDoc As FreeThreadedDOMDocument40
    Dim objICFClassDef As ICFClassDef

    Dim xmlAccountNode            As IXMLDOMNode, _
        xmlMortgageAccountNode    As IXMLDOMNode, _
        xmlIndemnityInsuranceNode As IXMLDOMNode
    Dim xmlThirdPartyGUIDNode As IXMLDOMNode, _
        xmlDirectoryGUIDNode  As IXMLDOMNode
    Dim objThirdPartyBO As ThirdPartyBO
    Dim xmlGUIDNode As IXMLDOMNode
    Dim strRequest As String
    Dim strAccountGUID As String
    
    ' PSC 22/07/02 BMIDS00006 - Start
    Dim xmlRelationships As IXMLDOMNodeList
    Dim xmlRelationship As IXMLDOMNode
    Dim xmlNode As IXMLDOMNode
    Dim xmlCopyNode As IXMLDOMNode
    ' PSC 22/07/02 BMIDS00006 - End
    
    Dim objIDOAssist As IDOAssist
    
    Dim objContext As ObjectContext
    Set objContext = GetObjectContext()
                    
    If objContext Is Nothing Then
        Set objIDOAssist = New DOAssist
        Set objThirdPartyBO = New ThirdPartyBO
    Else
        Set objIDOAssist = objContext.CreateInstance(gstrBASE_COMPONENT & ".DOAssist")
        Set objThirdPartyBO = objContext.CreateInstance(gstrTHIRDPARTY_COMPONENT & ".ThirdPartyBO")
    End If
    
    Set objICFClassDef = New CFClassDef
    strAccountGUID = CreateGUID()

    Set xmlMortgageAccountNode = vxmlTableElement
    Set xmlAccountNode = vxmlTableElement.selectSingleNode(".//ACCOUNT")
    Set xmlIndemnityInsuranceNode = vxmlTableElement.selectSingleNode(".//INDEMNITYINSURANCE")

    '
    ' Create ACCOUNT record
    '
    
    ' BG 15/06/00 SYS0928 - Start
    Set xmlGUIDNode = xmlAccountNode.selectSingleNode("./ACCOUNTGUID")
    
    If xmlGUIDNode Is Nothing Then
        Set xmlGUIDNode = xmlAccountNode.appendChild(xmlAccountNode.ownerDocument.createElement("ACCOUNTGUID"))
    End If
    
    xmlGUIDNode.Text = strAccountGUID
    ' BG 15/06/00 SYS0928 - End
    
    ' Maintain THIRDPARTY if necessary
    If (Trim$(objXmlAssist.GetTagValue(vxmlTableElement, "THIRDPARTY")) <> "") And _
       (Trim$(objXmlAssist.GetTagValue(vxmlTableElement, "DIRECTORYGUID")) = "") Then
        ' THIRDPARTY details exist in the XML, so maintain the appropriate THIRDPARTY record
        objTPAssist.MaintainThirdParty xmlAccountNode
    End If
    If Not (xmlAccountNode Is Nothing) Then
        objIDOAssist.Create xmlAccountNode, objICFClassDef.LoadAccountData
    End If
    
    '
    ' Create MORTGAGEACCOUNT record
    '
    ' BG 15/06/00 SYS0928 - Start
    Set xmlGUIDNode = xmlMortgageAccountNode.selectSingleNode("./ACCOUNTGUID")

    If xmlGUIDNode Is Nothing Then
        Set xmlGUIDNode = xmlMortgageAccountNode.appendChild(xmlMortgageAccountNode.ownerDocument.createElement("ACCOUNTGUID"))
    End If
    
    xmlGUIDNode.Text = strAccountGUID
    ' BG 15/06/00 SYS0928 - End
    
    objIDOAssist.Create xmlMortgageAccountNode, objICFClassDef.LoadMortgageAccountData
    
    ' PSC 22/07/2002 BMIDS00006 - Start
    ' Create the Account Relationships
    Set xmlNode = m_objXmlAssist.GetMandatoryNode(xmlMortgageAccountNode, "ACCOUNTRELATIONSHIPLIST")
    Set xmlCopyNode = xmlNode.cloneNode(True)
      
    Set xmlRelationships = xmlCopyNode.selectNodes("ACCOUNTRELATIONSHIP")
    If xmlRelationships.length = 0 Then
        m_objErrAssist.ThrowError TypeName(Me), strFunctionName, oeMissingPrimaryTag, "ACCOUNTRELATIONSHIPLIST/ACCOUNTRELATIONSHIP"
    End If
    
    For Each xmlRelationship In xmlRelationships
        Set xmlNode = m_objXmlAssist.GetNode(xmlRelationship, "ACCOUNTGUID", False)
        
        If xmlNode Is Nothing Then
            Set xmlNode = xmlRelationship.ownerDocument.createElement("ACCOUNTGUID")
            xmlRelationship.appendChild xmlNode
        End If
        
        xmlNode.Text = xmlGUIDNode.Text
    Next
    
    ICustomerFinancialDO_CreateAccountRelationships xmlCopyNode
    ' PSC 22/07/2002 BMIDS00006 - End
    
    '
    ' Create INDEMNITYINSURANCE record
    '
    ' BG 15/06/00 SYS0928 - Start
    If Not (xmlIndemnityInsuranceNode Is Nothing) Then
        Set xmlGUIDNode = xmlIndemnityInsuranceNode.selectSingleNode("./ACCOUNTGUID")

        If xmlGUIDNode Is Nothing Then
            Set xmlGUIDNode = xmlIndemnityInsuranceNode.appendChild(xmlIndemnityInsuranceNode.ownerDocument.createElement("ACCOUNTGUID"))
        End If
    
        xmlGUIDNode.Text = strAccountGUID
        objIDOAssist.Create xmlIndemnityInsuranceNode, objICFClassDef.LoadIndemnityInsuranceData
    End If
    ' BG 15/06/00 SYS0928 - End

    rxmlKeysNode.appendChild xmlGUIDNode
    
    Set objICFClassDef = Nothing
    Set xmlClassDefDoc = Nothing
    Set objIDOAssist = Nothing
    Set objErrAssist = Nothing
    Set xmlAccountNode = Nothing
    Set xmlMortgageAccountNode = Nothing
    Set xmlIndemnityInsuranceNode = Nothing
    Set objThirdPartyBO = Nothing
    Set xmlGUIDNode = Nothing
    Set xmlThirdPartyGUIDNode = Nothing
    Set xmlDirectoryGUIDNode = Nothing
    Set objXmlAssist = Nothing
    Set objTPAssist = Nothing
    
    ' PSC 22/07/02 BMIDS00006 - Start
    Set xmlRelationships = Nothing
    Set xmlRelationship = Nothing
    Set xmlNode = Nothing
    Set xmlCopyNode = Nothing
    ' PSC 22/07/02 BMIDS00006 - End
    
    If Not objContext Is Nothing Then
        objContext.SetComplete
    End If

    Exit Sub

CreateMortgageAccountVbErr:
    
    Set objICFClassDef = Nothing
    Set xmlClassDefDoc = Nothing
    Set objIDOAssist = Nothing
    Set xmlAccountNode = Nothing
    Set xmlMortgageAccountNode = Nothing
    Set xmlIndemnityInsuranceNode = Nothing
    Set objThirdPartyBO = Nothing
    Set xmlGUIDNode = Nothing
    Set xmlThirdPartyGUIDNode = Nothing
    Set xmlDirectoryGUIDNode = Nothing
    Set objXmlAssist = Nothing
    Set objTPAssist = Nothing
    
    ' PSC 22/07/02 BMIDS00006 - Start
    Set xmlRelationships = Nothing
    Set xmlRelationship = Nothing
    Set xmlNode = Nothing
    ' PSC 22/07/02 BMIDS00006 - End

    objErrAssist.AddToErrSource strFunctionName
    
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


Private Sub ICustomerFinancialDO_CreateMortgageLoan(ByVal vxmlTableElement As IXMLDOMElement, ByVal vxmlKeysNode As IXMLDOMElement)
' header ----------------------------------------------------------------------------------
' description:
'   CreateMortgageLoan an instance of the persistant data associated with this data object
'   for each set of data in the request
' pass:
'   vxmlxmlTableElement  xml element containing the request
' return:                n/a
'------------------------------------------------------------------------------------------
On Error GoTo CreateMortgageLoanVbErr

    Const strFunctionName As String = "ICustomerFinancialDO_CreateMortgageLoan"
   
    Dim objErrAssist  As New ErrAssist
    Dim xmlClassDefDoc As FreeThreadedDOMDocument40
    Dim objICFClassDef As ICFClassDef

    Dim xmlGUIDNode As IXMLDOMNode
    
    Dim objIDOAssist As IDOAssist
    
    Dim objContext As ObjectContext
    Set objContext = GetObjectContext()
                    
    If objContext Is Nothing Then
        Set objIDOAssist = New DOAssist
    Else
        Set objIDOAssist = objContext.CreateInstance(gstrBASE_COMPONENT & ".DOAssist")
    End If
    
    Set objICFClassDef = New CFClassDef
    Set xmlClassDefDoc = objICFClassDef.LoadMortgageLoanData()

    ' Generate the GUID for the new MORTGAGELOAN record
    Set xmlGUIDNode = vxmlTableElement.selectSingleNode("MORTGAGELOANGUID")
    If xmlGUIDNode Is Nothing Then
        Set xmlGUIDNode = vxmlTableElement.appendChild(vxmlTableElement.ownerDocument.createElement("MORTGAGELOANGUID"))
    End If
    xmlGUIDNode.Text = CreateGUID()

    objIDOAssist.Create vxmlTableElement, xmlClassDefDoc
    vxmlKeysNode.appendChild xmlGUIDNode
    
    Set objICFClassDef = Nothing
    Set xmlClassDefDoc = Nothing
    Set objIDOAssist = Nothing
    Set objErrAssist = Nothing
    Set xmlGUIDNode = Nothing

    If Not objContext Is Nothing Then
        objContext.SetComplete
    End If

    Exit Sub

CreateMortgageLoanVbErr:
    
    Set objICFClassDef = Nothing
    Set xmlClassDefDoc = Nothing
    Set objIDOAssist = Nothing
    Set xmlGUIDNode = Nothing
    
    objErrAssist.AddToErrSource strFunctionName
    
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

Private Sub ICustomerFinancialDO_CreateRegularOutgoings(ByVal vxmlTableElement As IXMLDOMElement)
' header ----------------------------------------------------------------------------------
' description:
'   CreateRegularOutgoings an instance of the persistant data associated with this data object
'   for each set of data in the request
' pass:
'   vxmlxmlTableElement  xml element containing the request
' return:                n/a
'------------------------------------------------------------------------------------------
On Error GoTo CreateRegularOutgoingsVbErr

    Const strFunctionName As String = "ICustomerFinancialDO_CreateRegularOutgoings"
   
    Dim xmlClassDefDoc      As FreeThreadedDOMDocument40
    Dim objICFClassDef      As ICFClassDef
    Dim xmlGUIDNode         As IXMLDOMNode
    'BMIDS00190
    Dim xmlCVROClassDefDoc  As FreeThreadedDOMDocument40
    Dim xmlRequest          As IXMLDOMElement
    Dim xmlCustomer         As IXMLDOMNode
    Dim xmlCopy             As IXMLDOMElement
    Dim strGuid             As String
    'BMIDS00190 End
    
    Set objICFClassDef = New CFClassDef
    Set xmlClassDefDoc = objICFClassDef.LoadRegularOutgoingsData()
    
    'BMIDS00190
    'Create a copy of vxmlTableElement and then remove the Customer Nodes
    Set xmlCopy = vxmlTableElement.cloneNode(True)
    
    For Each xmlCustomer In xmlCopy.selectNodes("CUSTOMERVERSIONREGULAROUTGOINGS")
        xmlCopy.removeChild xmlCustomer
    Next

    ' Generate the GUID for the new RegularOutgoings record
    Set xmlGUIDNode = xmlClassDefDoc.createElement("REGULAROUTGOINGSGUID")
    xmlGUIDNode.Text = CreateGUID()
    
    xmlCopy.appendChild xmlGUIDNode.cloneNode(True)
    
    m_objIDOAssist.Create xmlCopy, xmlClassDefDoc
    
    Set xmlCVROClassDefDoc = objICFClassDef.LoadCustVerRegularOutgoingsData()
    For Each xmlCustomer In vxmlTableElement.selectNodes("CUSTOMERVERSIONREGULAROUTGOINGS")
        xmlCustomer.appendChild xmlGUIDNode.cloneNode(True)
        m_objIDOAssist.Create xmlCustomer, xmlCVROClassDefDoc
    Next
    
    Set xmlCVROClassDefDoc = Nothing
    Set xmlRequest = Nothing
    Set xmlCustomer = Nothing
    Set xmlCopy = Nothing
    'BMIDS00190 End
    
    Set objICFClassDef = Nothing
    Set xmlClassDefDoc = Nothing
    Set xmlGUIDNode = Nothing

    If Not m_objContext Is Nothing Then
        m_objContext.SetComplete
    End If

    Exit Sub

CreateRegularOutgoingsVbErr:
    
    'BMIDS00190
    Set xmlCVROClassDefDoc = Nothing
    Set xmlRequest = Nothing
    Set xmlCustomer = Nothing
    Set xmlCopy = Nothing
    'BMIDS00190 End
    Set objICFClassDef = Nothing
    Set xmlClassDefDoc = Nothing
    Set xmlGUIDNode = Nothing
    
    m_objErrAssist.AddToErrSource strFunctionName
    
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

Private Sub ICustomerFinancialDO_DeleteAllMortgageLoans(ByVal vxmlTableElement As IXMLDOMElement)
' header ----------------------------------------------------------------------------------
' description:
'   Delete all instances of the persistant data associated with this
'   data object that match the key values specified
' pass:
'   vxmlxmlTableElement  xml element containing the request
' return:                n/a
'------------------------------------------------------------------------------------------
On Error GoTo DeleteAllMortgageLoansVbErr

    Const strFunctionName As String = "ICustomerFinancialDO_DeleteAllMortgageLoans"

    Dim objErrAssist As New ErrAssist
    
    Dim objIDOAssist As IDOAssist
    Dim xmlClassDefDoc As FreeThreadedDOMDocument40
    Dim objICFClassDef As ICFClassDef
    
    Dim objContext As ObjectContext
    Set objContext = GetObjectContext()
                            
    If objContext Is Nothing Then
        Set objIDOAssist = New DOAssist
    Else
        Set objIDOAssist = objContext.CreateInstance(gstrBASE_COMPONENT & ".DOAssist")
    End If
                
    Set objICFClassDef = New CFClassDef
    Set xmlClassDefDoc = objICFClassDef.LoadMortgageLoanData()
    objIDOAssist.DeleteAll vxmlTableElement, xmlClassDefDoc
 
    Set objICFClassDef = Nothing
    Set xmlClassDefDoc = Nothing
    Set objIDOAssist = Nothing
    Set objErrAssist = Nothing

    If Not objContext Is Nothing Then
        objContext.SetComplete
    End If

    Exit Sub

DeleteAllMortgageLoansVbErr:
    
    Set objICFClassDef = Nothing
    Set xmlClassDefDoc = Nothing
    Set objIDOAssist = Nothing
    
    objErrAssist.AddToErrSource strFunctionName
    
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

Private Sub ICustomerFinancialDO_DeleteArrearsHistory(ByVal vxmlTableElement As IXMLDOMElement)
' header ----------------------------------------------------------------------------------
' description:
'   DeleteArrearsHistory a single instance of the persistant data associated with this
'   data object
' pass:
'   vxmlxmlTableElement  xml element containing the request
' return:                n/a
'------------------------------------------------------------------------------------------
On Error GoTo DeleteArrearsHistoryVbErr

    Const strFunctionName As String = "ICustomerFinancialDO_DeleteArrearsHistory"

    Dim xmlClassDefDoc          As FreeThreadedDOMDocument40
    Dim objICFClassDef          As ICFClassDef
    'BMIDS00190
    Dim xmlAccountRelClassDef   As FreeThreadedDOMDocument40
    Dim xmlAccountClassDef      As FreeThreadedDOMDocument40
    Dim xmlOtherAAClassDef      As FreeThreadedDOMDocument40
    Dim xmlCustomer             As IXMLDOMNode
    Dim xmlOtherArrears         As IXMLDOMNode
    Dim xmlOtherArrearsTemp     As IXMLDOMNode
    Dim xmlRequest              As IXMLDOMNode
    Dim xmlAccountGuid          As IXMLDOMNode
    'BMIDS00190 End
                
    Set objICFClassDef = New CFClassDef
    Set xmlClassDefDoc = objICFClassDef.LoadArrearsHistoryData()
    
    'BMIDS00190
    Set xmlOtherArrearsTemp = vxmlTableElement.selectSingleNode("OTHERARREARSACCOUNT")
        
    If Not xmlOtherArrearsTemp Is Nothing Then
        Set xmlOtherArrears = xmlOtherArrearsTemp.cloneNode(True)
        vxmlTableElement.removeChild xmlOtherArrearsTemp
        
        Set xmlAccountGuid = xmlOtherArrears.selectSingleNode("ACCOUNTGUID")
        
        'Delete AccountRelationships records per customer
        Set xmlAccountRelClassDef = objICFClassDef.LoadAccountRelationshipData()
        For Each xmlCustomer In xmlOtherArrears.selectNodes("ACCOUNTRELATIONSHIP")
            xmlCustomer.appendChild xmlAccountGuid.cloneNode(True)
            m_objIDOAssist.Delete xmlCustomer, xmlAccountRelClassDef
            xmlOtherArrears.removeChild xmlCustomer
        Next
        
        'Delete OtherArrearsAccount record
        Set xmlOtherAAClassDef = objICFClassDef.LoadOtherArrearsAccountData()
        m_objIDOAssist.Delete xmlOtherArrears, xmlOtherAAClassDef
    End If
    'BMIDS00190 End
    
    'Delete ArrearsHistory record
    m_objIDOAssist.Delete vxmlTableElement, xmlClassDefDoc
    
    'BMIDS00190
    If Not xmlOtherArrears Is Nothing Then
        'Delete Account record last
        Set xmlAccountClassDef = objICFClassDef.LoadAccountData()
        
        Set xmlRequest = xmlClassDefDoc.createElement("ACCOUNT")
        xmlRequest.appendChild xmlAccountGuid.cloneNode(True)
        
        m_objIDOAssist.Delete xmlRequest, xmlAccountClassDef
    End If
    'BMIDS00190 End
    
    Set objICFClassDef = Nothing
    Set xmlClassDefDoc = Nothing
    'BMIDS00190
    Set xmlAccountRelClassDef = Nothing
    Set xmlAccountClassDef = Nothing
    Set xmlOtherAAClassDef = Nothing
    Set xmlCustomer = Nothing
    Set xmlOtherArrears = Nothing
    Set xmlOtherArrearsTemp = Nothing
    Set xmlRequest = Nothing
    Set xmlAccountGuid = Nothing
    'BMIDS00190 End
    
    If Not m_objContext Is Nothing Then
        m_objContext.SetComplete
    End If
    
    Exit Sub

DeleteArrearsHistoryVbErr:

    Set objICFClassDef = Nothing
    Set xmlClassDefDoc = Nothing
    'BMIDS00190
    Set xmlAccountRelClassDef = Nothing
    Set xmlAccountClassDef = Nothing
    Set xmlOtherAAClassDef = Nothing
    Set xmlCustomer = Nothing
    Set xmlOtherArrears = Nothing
    Set xmlOtherArrearsTemp = Nothing
    Set xmlRequest = Nothing
    Set xmlAccountGuid = Nothing
    'BMIDS00190 End
    
    m_objErrAssist.AddToErrSource strFunctionName
    
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

Private Sub ICustomerFinancialDO_DeleteBankCard(ByVal vxmlTableElement As IXMLDOMElement)
' header ----------------------------------------------------------------------------------
' description:
'   DeleteBankCard a single instance of the persistant data associated with this
'   data object
' pass:
'   vxmlxmlTableElement  xml element containing the request
' return:                n/a
'------------------------------------------------------------------------------------------
On Error GoTo DeleteBankCardVbErr

    Const strFunctionName As String = "ICustomerFinancialDO_DeleteBankCard"

    Dim objErrAssist As New ErrAssist
    
    Dim objContext As ObjectContext
    Set objContext = GetObjectContext()
            
    Dim objIDOAssist As IDOAssist
    Dim xmlClassDefDoc As FreeThreadedDOMDocument40
    Dim objICFClassDef As ICFClassDef
    
    If objContext Is Nothing Then
        Set objIDOAssist = New DOAssist
    Else
        Set objIDOAssist = objContext.CreateInstance(gstrBASE_COMPONENT & ".DOAssist")
    End If
            
    Set objICFClassDef = New CFClassDef
    Set xmlClassDefDoc = objICFClassDef.LoadBankCreditCardData
    objIDOAssist.Delete vxmlTableElement, xmlClassDefDoc
    
    Set objICFClassDef = Nothing
    Set xmlClassDefDoc = Nothing
    Set objIDOAssist = Nothing
    Set objErrAssist = Nothing
    
    If Not objContext Is Nothing Then
        objContext.SetComplete
    End If
    
    Exit Sub

DeleteBankCardVbErr:

    Set objICFClassDef = Nothing
    Set xmlClassDefDoc = Nothing
    Set objIDOAssist = Nothing
    
    objErrAssist.AddToErrSource strFunctionName
    
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

Private Sub ICustomerFinancialDO_DeleteBankruptcyHistory(ByVal vxmlTableElement As IXMLDOMElement)
' header ----------------------------------------------------------------------------------
' description:
'   DeleteBankruptcyHistory a single instance of the persistant data associated with this
'   data object
' pass:
'   vxmlxmlTableElement  xml element containing the request
' return:                n/a
'------------------------------------------------------------------------------------------
On Error GoTo DeleteBankruptcyHistoryVbErr

    Const strFunctionName As String = "ICustomerFinancialDO_DeleteBankruptcyHistory"
   
    Dim xmlClassDefDoc      As FreeThreadedDOMDocument40
    Dim objICFClassDef      As ICFClassDef
    'BMIDS00190
    Dim xmlCVBHClassDefDoc  As FreeThreadedDOMDocument40
    Dim xmlCustomer         As IXMLDOMNode
    'BMIDS00190 End
               
    Set objICFClassDef = New CFClassDef
    Set xmlClassDefDoc = objICFClassDef.LoadBankruptcyHistoryData()
    
    'BMIDS00190 Remove all customer links
    Set xmlCVBHClassDefDoc = objICFClassDef.LoadCustVerBankruptcyHistoryData()
    For Each xmlCustomer In vxmlTableElement.selectNodes("CUSTOMERVERSIONBANKRUPTCYHISTORY")
        xmlCustomer.appendChild vxmlTableElement.selectSingleNode("BANKRUPTCYHISTORYGUID").cloneNode(True)
        m_objIDOAssist.Delete xmlCustomer, xmlCVBHClassDefDoc
        vxmlTableElement.removeChild xmlCustomer
    Next
    'BMIDS00190 End
    
    m_objIDOAssist.Delete vxmlTableElement, xmlClassDefDoc
    
    Set objICFClassDef = Nothing
    Set xmlClassDefDoc = Nothing
    'BMIDS00190
    Set xmlCVBHClassDefDoc = Nothing
    Set xmlCustomer = Nothing
    'BMIDS00190 End
    
    If Not m_objContext Is Nothing Then
        m_objContext.SetComplete
    End If
    
    Exit Sub

DeleteBankruptcyHistoryVbErr:

    Set objICFClassDef = Nothing
    Set xmlClassDefDoc = Nothing
    'BMIDS00190
    Set xmlCVBHClassDefDoc = Nothing
    Set xmlCustomer = Nothing
    'BMIDS00190 End
    
    m_objErrAssist.AddToErrSource strFunctionName
    
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

Private Sub ICustomerFinancialDO_DeleteCCJHistory(ByVal vxmlTableElement As IXMLDOMElement)
' header ----------------------------------------------------------------------------------
' description:
'   DeleteCCJHistory a single instance of the persistant data associated with this
'   data object
' pass:
'   vxmlxmlTableElement  xml element containing the request
' return:                n/a
'------------------------------------------------------------------------------------------
On Error GoTo DeleteCCJHistoryVbErr

    Const strFunctionName As String = "ICustomerFinancialDO_DeleteCCJHistory"
    
    Dim xmlClassDefDoc      As FreeThreadedDOMDocument40
    Dim objICFClassDef      As ICFClassDef
    'BMIDS00190
    Dim xmlCVCHClassDefDoc  As FreeThreadedDOMDocument40
    Dim xmlCustomer         As IXMLDOMNode
    'BMIDS00190 End
    
    Set objICFClassDef = New CFClassDef
    Set xmlClassDefDoc = objICFClassDef.LoadCCJHistoryData()
    
    'BMIDS00190 Remove all customer links
    Set xmlCVCHClassDefDoc = objICFClassDef.LoadCustVerCCJHistoryData()
    For Each xmlCustomer In vxmlTableElement.selectNodes("CUSTOMERVERSIONCCJHISTORY")
        xmlCustomer.appendChild vxmlTableElement.selectSingleNode("CCJHISTORYGUID").cloneNode(True)
        m_objIDOAssist.Delete xmlCustomer, xmlCVCHClassDefDoc
        vxmlTableElement.removeChild xmlCustomer
    Next
    'BMIDS00190 End
    
    m_objIDOAssist.Delete vxmlTableElement, xmlClassDefDoc
    
    Set objICFClassDef = Nothing
    Set xmlClassDefDoc = Nothing
    'BMIDS00190
    Set xmlCVCHClassDefDoc = Nothing
    Set xmlCustomer = Nothing
    'BMIDS00190 End
    
    If Not m_objContext Is Nothing Then
        m_objContext.SetComplete
    End If
    
    Exit Sub

DeleteCCJHistoryVbErr:

    Set objICFClassDef = Nothing
    Set xmlClassDefDoc = Nothing
    'BMIDS00190
    Set xmlCVCHClassDefDoc = Nothing
    Set xmlCustomer = Nothing
    'BMIDS00190 End
    
    m_objErrAssist.AddToErrSource strFunctionName
    
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

Private Sub ICustomerFinancialDO_DeleteDeclinedMortgage(ByVal vxmlTableElement As IXMLDOMElement)
' header ----------------------------------------------------------------------------------
' description:
'   DeleteDeclinedMortgage a single instance of the persistant data associated with this
'   data object
' pass:
'   vxmlxmlTableElement  xml element containing the request
' return:                n/a
'------------------------------------------------------------------------------------------
On Error GoTo DeleteDeclinedMortgageVbErr

    Const strFunctionName As String = "ICustomerFinancialDO_DeleteDeclinedMortgage"

    Dim objErrAssist As New ErrAssist
    
    Dim objContext As ObjectContext
    Set objContext = GetObjectContext()
            
    Dim objIDOAssist As IDOAssist
    Dim xmlClassDefDoc As FreeThreadedDOMDocument40
    Dim objICFClassDef As ICFClassDef
    
    If objContext Is Nothing Then
        Set objIDOAssist = New DOAssist
    Else
        Set objIDOAssist = objContext.CreateInstance(gstrBASE_COMPONENT & ".DOAssist")
    End If
            
    Set objICFClassDef = New CFClassDef
    Set xmlClassDefDoc = objICFClassDef.LoadDeclinedMortgageData
    objIDOAssist.Delete vxmlTableElement, xmlClassDefDoc
    
    Set objICFClassDef = Nothing
    Set xmlClassDefDoc = Nothing
    Set objIDOAssist = Nothing
    Set objErrAssist = Nothing
    
    If Not objContext Is Nothing Then
        objContext.SetComplete
    End If
    
    Exit Sub

DeleteDeclinedMortgageVbErr:

    Set objICFClassDef = Nothing
    Set xmlClassDefDoc = Nothing
    Set objIDOAssist = Nothing
    
    objErrAssist.AddToErrSource strFunctionName
    
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

Private Sub ICustomerFinancialDO_DeleteGroupConnection(ByVal vxmlTableElement As IXMLDOMElement)
' header ----------------------------------------------------------------------------------
' description:
'   DeleteGroupConnection a single instance of the persistant data associated with this
'   data object
' pass:
'   vxmlxmlTableElement  xml element containing the request
' return:                n/a
'------------------------------------------------------------------------------------------
On Error GoTo DeleteGroupConnectionVbErr

    Const strFunctionName As String = "ICustomerFinancialDO_DeleteGroupConnection"

    Dim objErrAssist As New ErrAssist
    
    Dim objContext As ObjectContext
    Set objContext = GetObjectContext()
            
    Dim objIDOAssist As IDOAssist
    Dim xmlClassDefDoc As FreeThreadedDOMDocument40
    Dim objICFClassDef As ICFClassDef
    
    If objContext Is Nothing Then
        Set objIDOAssist = New DOAssist
    Else
        Set objIDOAssist = objContext.CreateInstance(gstrBASE_COMPONENT & ".DOAssist")
    End If
            
    Set objICFClassDef = New CFClassDef
    Set xmlClassDefDoc = objICFClassDef.LoadGroupConnectionData()
    objIDOAssist.Delete vxmlTableElement, xmlClassDefDoc
    
    Set objICFClassDef = Nothing
    Set xmlClassDefDoc = Nothing
    Set objIDOAssist = Nothing
    Set objErrAssist = Nothing
    
    If Not objContext Is Nothing Then
        objContext.SetComplete
    End If
    
    Exit Sub

DeleteGroupConnectionVbErr:

    Set objICFClassDef = Nothing
    Set xmlClassDefDoc = Nothing
    Set objIDOAssist = Nothing
    
    objErrAssist.AddToErrSource strFunctionName
    
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

Private Sub ICustomerFinancialDO_DeleteLiability(ByVal vxmlTableElement As IXMLDOMElement, _
                                                 Optional ByVal vblnSimpleDelete As Boolean = False)
' header ----------------------------------------------------------------------------------
' description:
'   DeleteLiability a single instance of the persistant data associated with this
'   data object
' pass:
'   vxmlxmlTableElement  xml element containing the request
' return:                n/a
'------------------------------------------------------------------------------------------
On Error GoTo DeleteLiabilityVbErr

    Const strFunctionName As String = "ICustomerFinancialDO_DeleteLiability"
    
    
    Dim objErrAssist As New ErrAssist
    
    Dim objContext As ObjectContext
    Set objContext = GetObjectContext()
            
    Dim objIADOAssist As IADOAssist
    Dim Cmd As ADODB.Command
    Dim objSQLAssist As New SQLAssist

    Dim xmlAccountGUIDNode     As IXMLDOMNode
    Dim strAccountGUID As String
    
    If objContext Is Nothing Then
        Set objIADOAssist = New DOAssist
    Else
        Set objIADOAssist = objContext.CreateInstance(gstrBASE_COMPONENT & ".ADOAssist")
    End If

    Set xmlAccountGUIDNode = vxmlTableElement.selectSingleNode(".//ACCOUNTGUID")

    If (xmlAccountGUIDNode Is Nothing) Then
        objErrAssist.ThrowError TypeName(Me), strFunctionName, oeMissingPrimaryTag
    End If

    strAccountGUID = xmlAccountGUIDNode.Text

    Set Cmd = New ADODB.Command
    With Cmd
        .ActiveConnection = objIADOAssist.GetConnStr
        ' DM 09/09/01 SYS2572
        #If GENERIC_SQL Then
            
            Dim rbyteAccountGuid() As Byte
            
            ' Convert the GUID string to a binary
            rbyteAccountGuid() = objSQLAssist.GuidStringToByteArray(strAccountGUID)

            .CommandText = "DELETELOANSLIABILITIES_GS"
            .Parameters.Append .CreateParameter("rawAccountGuid", adVarBinary, adParamInput, UBound(rbyteAccountGuid) + 1, rbyteAccountGuid)
            .Parameters.Append .CreateParameter("bSaveThirdParty", adInteger, adParamInput, 1, vblnSimpleDelete)
            
        #Else
            .CommandText = "DELETELOANSLIABILITIES_GS(" & _
                            objSQLAssist.FormatGuid(strAccountGUID) & "," & _
                            vblnSimpleDelete & ")"
        #End If
        .Execute , , adCmdStoredProc
    End With
    
    Set objIADOAssist = Nothing
    Set Cmd = Nothing
    Set objErrAssist = Nothing
    Set objSQLAssist = Nothing
    Set xmlAccountGUIDNode = Nothing
    
    If Not objContext Is Nothing Then
        objContext.SetComplete
    End If
    Exit Sub

DeleteLiabilityVbErr:

    Set objIADOAssist = Nothing
    Set Cmd = Nothing
    Set objSQLAssist = Nothing
    Set xmlAccountGUIDNode = Nothing
    
    objErrAssist.AddToErrSource strFunctionName
    
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

Private Sub ICustomerFinancialDO_DeleteLifeProduct(ByVal vxmlTableElement As IXMLDOMElement, _
                                                   Optional ByVal vblnSimpleDelete As Boolean = False)
' header ----------------------------------------------------------------------------------
' description:
'   DeleteLifeProduct a single instance of the persistant data associated with this
'   data object
' pass:
'   vxmlxmlTableElement  xml element containing the request
' return:                n/a
'------------------------------------------------------------------------------------------
On Error GoTo DeleteLifeProductVbErr

    Const strFunctionName As String = "ICustomerFinancialDO_DeleteLifeProduct"

    Dim objErrAssist As New ErrAssist
    
    Dim objContext As ObjectContext
    Set objContext = GetObjectContext()
            
    Dim objIDOAssist As IDOAssist
    Dim objIADOAssist As IADOAssist
    Dim Cmd As ADODB.Command
    Dim objSQLAssist As New SQLAssist
    Dim objICFClassDef As ICFClassDef

    Dim xmlGUIDNode            As IXMLDOMNode, _
        xmlCustomerNumberNode  As IXMLDOMNode, _
        xmlCustomerVersionNode As IXMLDOMNode

    If vblnSimpleDelete Then
        ' Just delete the item from the policy relationship table
        If objContext Is Nothing Then
            Set objIDOAssist = New DOAssist
        Else
            Set objIDOAssist = objContext.CreateInstance(gstrBASE_COMPONENT & ".DOAssist")
        End If

        Set objICFClassDef = New CFClassDef
    Else
        ' Delete a record from ALL the tables in the LIFEPRODUCT join
        If objContext Is Nothing Then
            Set objIADOAssist = New ADOAssist
        Else
            Set objIADOAssist = objContext.CreateInstance(gstrBASE_COMPONENT & ".ADOAssist")
        End If

        Set xmlGUIDNode = vxmlTableElement.selectSingleNode(".//ACCOUNTGUID")
        Set xmlCustomerNumberNode = vxmlTableElement.selectSingleNode(".//CUSTOMERNUMBER")
        Set xmlCustomerVersionNode = vxmlTableElement.selectSingleNode(".//CUSTOMERVERSIONNUMBER")
        If (xmlGUIDNode Is Nothing) Or _
           (xmlCustomerNumberNode Is Nothing) Or _
           (xmlCustomerVersionNode Is Nothing) Then
            objErrAssist.ThrowError TypeName(Me), strFunctionName, oeMissingPrimaryTag
        End If
        
        Set Cmd = New ADODB.Command
        With Cmd
            .ActiveConnection = objIADOAssist.GetConnStr
            ' DM 09/09/01 SYS2572
            #If GENERIC_SQL Then
                ' Need to change proc name
                ' need to format the guids.
                Dim rbyteAccountGuid() As Byte
                rbyteAccountGuid() = objSQLAssist.GuidStringToByteArray(xmlGUIDNode.Text)
                .CommandText = "DELETELIFEPRODUCT_GS"
                .Parameters.Append .CreateParameter("strCustomerNumber", adVarChar, adParamInput, Len(xmlCustomerNumberNode.Text), xmlCustomerNumberNode.Text)
                .Parameters.Append .CreateParameter("intCUSTOMERVERSIONNUMBER", adInteger, adParamInput, , CInt(xmlCustomerVersionNode.Text))
                .Parameters.Append .CreateParameter("rawACCOUNTGUID", adVarBinary, adParamInput, UBound(rbyteAccountGuid) + 1, rbyteAccountGuid)
            #Else
                .CommandText = "DELETELIFEPRODUCT(" & _
                                objSQLAssist.FormatString(xmlCustomerNumberNode.Text) & "," & _
                                xmlCustomerVersionNode.Text & "," & _
                                objSQLAssist.FormatGuid(xmlGUIDNode.Text) & ")"
            #End If
            .Execute , , adCmdStoredProc
        End With
    End If

    
    Set objIADOAssist = Nothing
    Set objIDOAssist = Nothing
    Set Cmd = Nothing
    Set objErrAssist = Nothing
    Set objSQLAssist = Nothing
    Set xmlGUIDNode = Nothing
    Set xmlCustomerNumberNode = Nothing
    Set xmlCustomerVersionNode = Nothing
    Set objICFClassDef = Nothing
    
    If Not objContext Is Nothing Then
        objContext.SetComplete
    End If
    
    Exit Sub

DeleteLifeProductVbErr:

    Set objIADOAssist = Nothing
    Set objIDOAssist = Nothing
    Set Cmd = Nothing
    Set objSQLAssist = Nothing
    Set xmlGUIDNode = Nothing
    Set xmlCustomerNumberNode = Nothing
    Set xmlCustomerVersionNode = Nothing
    Set objICFClassDef = Nothing
    
    objErrAssist.AddToErrSource strFunctionName
    
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

Private Sub ICustomerFinancialDO_DeleteMortgageAccount(ByVal vxmlTableElement As IXMLDOMElement)
' header ----------------------------------------------------------------------------------
' description:
'   DeleteMortgageAccount a single instance of the persistant data associated with this
'   data object
' pass:
'   vxmlxmlTableElement  xml element containing the request
' return:                n/a
'------------------------------------------------------------------------------------------
On Error GoTo DeleteMortgageAccountVbErr

    Const strFunctionName As String = "ICustomerFinancialDO_DeleteMortgageAccount"

    Dim objErrAssist As New ErrAssist
    
    Dim objContext As ObjectContext
    Set objContext = GetObjectContext()
            
    Dim objIADOAssist As IADOAssist
    Dim Cmd As ADODB.Command
    Dim objSQLAssist As New SQLAssist

    Dim xmlAccountGUIDNode     As IXMLDOMNode, _
        xmlThirdPartyGUIDNode  As IXMLDOMNode
    Dim strAccountGUID As String, _
        strThirdPartyGUID As String
    
    If objContext Is Nothing Then
        Set objIADOAssist = New DOAssist
    Else
        Set objIADOAssist = objContext.CreateInstance(gstrBASE_COMPONENT & ".ADOAssist")
    End If

    Set xmlAccountGUIDNode = vxmlTableElement.selectSingleNode(".//ACCOUNTGUID")
    Set xmlThirdPartyGUIDNode = vxmlTableElement.selectSingleNode(".//THIRDPARTYGUID")

    If (xmlAccountGUIDNode Is Nothing) Then
        objErrAssist.ThrowError TypeName(Me), strFunctionName, oeMissingPrimaryTag
    End If

    strAccountGUID = xmlAccountGUIDNode.Text
    If Not (xmlThirdPartyGUIDNode Is Nothing) Then
        strThirdPartyGUID = xmlThirdPartyGUIDNode.Text
    End If

    Set Cmd = New ADODB.Command
    With Cmd
        .ActiveConnection = objIADOAssist.GetConnStr
        ' DM 09/09/01 SYS2572
        #If GENERIC_SQL Then
            
            Dim rbyteAccountGuid() As Byte
            Dim rbyteThirdPartyGuid() As Byte
            
            ' Convert the GUID string to a binary
            rbyteAccountGuid() = objSQLAssist.GuidStringToByteArray(strAccountGUID)

            .CommandText = "DELETEMORTGAGEACCOUNT_GS"
            .Parameters.Append .CreateParameter("rawAccountGuid", adVarBinary, adParamInput, UBound(rbyteAccountGuid) + 1, rbyteAccountGuid)
            If Not strThirdPartyGUID = "" Then
                rbyteThirdPartyGuid() = objSQLAssist.GuidStringToByteArray(strThirdPartyGUID)
                .Parameters.Append .CreateParameter("rawThirdPartyGuid", adVarBinary, adParamInput, UBound(rbyteThirdPartyGuid) + 1, rbyteThirdPartyGuid)
            Else
                ' need to create a parameter here but set it to null.
                .Parameters.Append .CreateParameter("rawThirdPartyGuid", adVarBinary, adParamInput, 1, Null)
            End If
            
            
            
        #Else
            .CommandText = "DELETEMORTGAGEACCOUNT(" & _
                            objSQLAssist.FormatGuid(strAccountGUID) & "," & _
                            objSQLAssist.FormatGuid(strThirdPartyGUID) & ")"
        #End If
        .Execute , , adCmdStoredProc
    End With
    
    Set objIADOAssist = Nothing
    Set Cmd = Nothing
    Set objErrAssist = Nothing
    Set objSQLAssist = Nothing
    Set xmlAccountGUIDNode = Nothing
    Set xmlThirdPartyGUIDNode = Nothing
    
    If Not objContext Is Nothing Then
        objContext.SetComplete
    End If
    
    Exit Sub

DeleteMortgageAccountVbErr:

    Set objIADOAssist = Nothing
    Set Cmd = Nothing
    Set objSQLAssist = Nothing
    Set xmlAccountGUIDNode = Nothing
    Set xmlThirdPartyGUIDNode = Nothing
    
    objErrAssist.AddToErrSource strFunctionName
    
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

Private Sub ICustomerFinancialDO_DeleteMortgageLoan(ByVal vxmlTableElement As IXMLDOMElement)
' header ----------------------------------------------------------------------------------
' description:
'   DeleteMortgageLoan a single instance of the persistant data associated with this
'   data object
' pass:
'   vxmlxmlTableElement  xml element containing the request
' return:                n/a
'------------------------------------------------------------------------------------------
On Error GoTo DeleteMortgageLoanVbErr

    Const strFunctionName As String = "ICustomerFinancialDO_DeleteMortgageLoan"

    Dim objErrAssist As New ErrAssist
    
    Dim objContext As ObjectContext
    Set objContext = GetObjectContext()
            
    Dim objIDOAssist As IDOAssist
    Dim xmlClassDefDoc As FreeThreadedDOMDocument40
    Dim objICFClassDef As ICFClassDef
    
    If objContext Is Nothing Then
        Set objIDOAssist = New DOAssist
    Else
        Set objIDOAssist = objContext.CreateInstance(gstrBASE_COMPONENT & ".DOAssist")
    End If
            
    Set objICFClassDef = New CFClassDef
    Set xmlClassDefDoc = objICFClassDef.LoadMortgageLoanData()
    objIDOAssist.Delete vxmlTableElement, xmlClassDefDoc
    
    Set objICFClassDef = Nothing
    Set xmlClassDefDoc = Nothing
    Set objIDOAssist = Nothing
    Set objErrAssist = Nothing
    
    If Not objContext Is Nothing Then
        objContext.SetComplete
    End If
    
    Exit Sub

DeleteMortgageLoanVbErr:

    Set objICFClassDef = Nothing
    Set xmlClassDefDoc = Nothing
    Set objIDOAssist = Nothing
    
    objErrAssist.AddToErrSource strFunctionName
    
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

Private Sub ICustomerFinancialDO_DeleteRegularOutgoings(ByVal vxmlTableElement As IXMLDOMElement)
' header ----------------------------------------------------------------------------------
' description:
'   DeleteMortgageLoan a single instance of the persistant data associated with this
'   data object
' pass:
'   vxmlxmlTableElement  xml element containing the request
' return:                n/a
'------------------------------------------------------------------------------------------
On Error GoTo DeleteRegularOutgoingsVbErr

    Const strFunctionName As String = "ICustomerFinancialDO_DeleteRegularOutgoings"
    
    Dim xmlClassDefDoc      As FreeThreadedDOMDocument40
    Dim objICFClassDef      As ICFClassDef
    'BMIDS00190
    Dim xmlCVROClassDefDoc  As FreeThreadedDOMDocument40
    Dim xmlCustomer         As IXMLDOMNode
    'BMIDS00190 End
    
    Set objICFClassDef = New CFClassDef
    Set xmlClassDefDoc = objICFClassDef.LoadRegularOutgoingsData()
    
    'BMIDS00190 Remove all customer links
    Set xmlCVROClassDefDoc = objICFClassDef.LoadCustVerRegularOutgoingsData()
    For Each xmlCustomer In vxmlTableElement.selectNodes("CUSTOMERVERSIONREGULAROUTGOINGS")
        xmlCustomer.appendChild vxmlTableElement.selectSingleNode("REGULAROUTGOINGSGUID").cloneNode(True)
        m_objIDOAssist.Delete xmlCustomer, xmlCVROClassDefDoc
        vxmlTableElement.removeChild xmlCustomer
    Next
    'BMIDS00190 End
    
    m_objIDOAssist.Delete vxmlTableElement, xmlClassDefDoc
    
    Set objICFClassDef = Nothing
    Set xmlClassDefDoc = Nothing
    'BMIDS00190
    Set xmlCVROClassDefDoc = Nothing
    Set xmlCustomer = Nothing
    'BMIDS00190 End
    
    If Not m_objContext Is Nothing Then
        m_objContext.SetComplete
    End If
    
    Exit Sub

DeleteRegularOutgoingsVbErr:

    Set objICFClassDef = Nothing
    Set xmlClassDefDoc = Nothing
    'BMIDS00190
    Set xmlCVROClassDefDoc = Nothing
    Set xmlCustomer = Nothing
    'BMIDS00190 End
    
    m_objErrAssist.AddToErrSource strFunctionName
    
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

Private Function ICustomerFinancialDO_FindArrearsHistoryList(ByVal vxmlTableElement As IXMLDOMElement) As IXMLDOMNode
' header ----------------------------------------------------------------------------------
' description:
'   Get the data for all instances of the persistant data associated with
'   this data object for the values supplied
' pass:
'   vxmlxmlTableElement  xml element containing the request
' return:                xml node containing retrieved data
'------------------------------------------------------------------------------------------
On Error GoTo FindArrearsHistoryListVbErr

    Const strFunctionName As String = "ICustomerFinancialDO_FindArrearsHistoryList"
   
    Dim xmlClassDefDoc As FreeThreadedDOMDocument40
    Dim objICFClassDef As ICFClassDef
    Dim xmlReturn As IXMLDOMNode
    Dim xmlCustomerNoList As IXMLDOMNodeList
    
    'BMIDS00190
    Dim xmlCFAHList         As IXMLDOMNode
    Dim xmlArrearsList      As IXMLDOMNode
    Dim xmlArrears          As IXMLDOMNode
    Dim xmlCFAH             As IXMLDOMNode
    Dim strGuid             As String
    Dim xmlAccountRel       As IXMLDOMNode
    Dim xmlDescription      As IXMLDOMNode
    Dim xmlOtherArrears     As IXMLDOMNode
    'BMIDS00190 End
    
    Set objICFClassDef = New CFClassDef
    Set xmlClassDefDoc = objICFClassDef.LoadCFArrearsHistoryData()
    
    'BMIDS00190
    Set xmlArrearsList = xmlClassDefDoc.createElement("ARREARSHISTORYLIST")
    Set xmlCFAHList = m_objIDOAssist.FindListMultiple(vxmlTableElement, xmlClassDefDoc, "ACCOUNTGUID, CUSTOMERNUMBER", "ARREARSHISTORY")
    
    For Each xmlCFAH In xmlCFAHList.selectNodes("CF_ARREARSHISTORY")
        ' RF 19/07/04 BMIDS788 Start - remove recursion in call to ChangeNodeName
        m_objXmlAssist.ChangeNodeName xmlCFAH, "CF_ARREARSHISTORY", "ARREARSHISTORY", False
        ' RF 19/07/04 BMIDS788 End
        
        strGuid = m_objXmlAssist.GetMandatoryElementText(xmlCFAH, "ARREARSHISTORYGUID")
        Set xmlArrears = xmlArrearsList.selectSingleNode("ARREARSHISTORY[ARREARSHISTORYGUID[.='" & strGuid & "']]")
        If xmlArrears Is Nothing Then
            ' Create new arrears history node
            Set xmlArrears = xmlCFAH.cloneNode(True)
            
            Set xmlOtherArrears = xmlArrears.selectSingleNode("OTHERARREARSACCOUNT")
            Set xmlDescription = xmlOtherArrears.selectSingleNode("DESCRIPTION")
            If Trim(Len(xmlDescription.Text)) = 0 Then
                xmlArrears.removeChild xmlOtherArrears
            End If
            
            xmlArrearsList.appendChild xmlArrears
        Else
            xmlArrears.appendChild xmlCFAH.selectSingleNode("ACCOUNTRELATIONSHIP")
        End If
    Next
    
    Set xmlReturn = xmlArrearsList
    'BMIDS00190 End
                
'    ' PSC 12/08/2002 BMIDS00006 - Start
'    Set xmlCustomerNoList = vxmlTableElement.selectNodes("CUSTOMERNUMBER")
'
'    If xmlCustomerNoList.length > 0 Then
'        OrderListByApplicant vxmlTableElement, xmlReturn, "ARREARSHISTORY"
'    End If
'    ' PSC 12/08/2002 BMIDS00006 - End
    
    Set ICustomerFinancialDO_FindArrearsHistoryList = xmlReturn

    Set objICFClassDef = Nothing
    Set xmlClassDefDoc = Nothing
    Set xmlReturn = Nothing
    Set xmlCustomerNoList = Nothing
    
    'BMIDS00190
    Set xmlCFAHList = Nothing
    Set xmlArrearsList = Nothing
    Set xmlArrears = Nothing
    Set xmlCFAH = Nothing
    Set xmlAccountRel = Nothing
    Set xmlDescription = Nothing
    Set xmlOtherArrears = Nothing
    'BMIDS00190 End
    
    If Not m_objContext Is Nothing Then
        m_objContext.SetComplete
    End If
    
    Exit Function

FindArrearsHistoryListVbErr:
    
    Set objICFClassDef = Nothing
    Set xmlClassDefDoc = Nothing
    Set xmlReturn = Nothing
    Set xmlCustomerNoList = Nothing
    
    'BMIDS00190
    Set xmlCFAHList = Nothing
    Set xmlArrearsList = Nothing
    Set xmlArrears = Nothing
    Set xmlCFAH = Nothing
    Set xmlAccountRel = Nothing
    Set xmlDescription = Nothing
    Set xmlOtherArrears = Nothing
    'BMIDS00190 End
        
    m_objErrAssist.AddToErrSource strFunctionName
    
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

Private Function ICustomerFinancialDO_FindBankCardList(ByVal vxmlTableElement As IXMLDOMElement) As IXMLDOMNode
' header ----------------------------------------------------------------------------------
' description:
'   Get the data for all instances of the persistant data associated with
'   this data object for the values supplied
' pass:
'   vxmlxmlTableElement  xml element containing the request
' return:                xml node containing retrieved data
'------------------------------------------------------------------------------------------
On Error GoTo FindBankCardListVbErr

    Const strFunctionName As String = "ICustomerFinancialDO_FindBankCardList"

    Dim objErrAssist As New ErrAssist
    
    Dim objIDOAssist As IDOAssist
    Dim xmlClassDefDoc As FreeThreadedDOMDocument40
    Dim objICFClassDef As ICFClassDef
    Dim xmlReturn As IXMLDOMNode
    
    Dim objContext As ObjectContext
    Set objContext = GetObjectContext()
         
    If objContext Is Nothing Then
        Set objIDOAssist = New DOAssist
    Else
        Set objIDOAssist = objContext.CreateInstance(gstrBASE_COMPONENT & ".DOAssist")
    End If
     
    Set objICFClassDef = New CFClassDef
    Set xmlClassDefDoc = objICFClassDef.LoadBankCreditCardData
    Set xmlReturn = objIDOAssist.FindListMultiple(vxmlTableElement, xmlClassDefDoc, _
            vstrOrderByField:="CUSTOMERNUMBER,CUSTOMERVERSIONNUMBER,SEQUENCENUMBER")

    OrderListByApplicant vxmlTableElement, xmlReturn, "BANKCREDITCARD"

    Set ICustomerFinancialDO_FindBankCardList = xmlReturn

    Set objICFClassDef = Nothing
    Set xmlClassDefDoc = Nothing
    Set objErrAssist = Nothing
    Set objIDOAssist = Nothing
    Set xmlReturn = Nothing
    
    'If Not objContext Is Nothing Then
    '    objContext.SetComplete
    'End If
    
    Exit Function

FindBankCardListVbErr:
    
    Set objICFClassDef = Nothing
    Set xmlClassDefDoc = Nothing
    Set objIDOAssist = Nothing
    Set xmlReturn = Nothing
    
    objErrAssist.AddToErrSource strFunctionName
    
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

Private Function ICustomerFinancialDO_FindBankruptcyHistoryList(ByVal vxmlTableElement As IXMLDOMElement) As IXMLDOMNode
' header ----------------------------------------------------------------------------------
' description:
'   Get the data for all instances of the persistant data associated with
'   this data object for the values supplied
' pass:
'   vxmlxmlTableElement  xml element containing the request
' return:                xml node containing retrieved data
'------------------------------------------------------------------------------------------
On Error GoTo FindBankruptcyHistoryListVbErr

    Const strFunctionName As String = "ICustomerFinancialDO_FindBankruptcyHistoryList"
   
    Dim xmlClassDefDoc      As FreeThreadedDOMDocument40
    Dim objICFClassDef      As ICFClassDef
    Dim xmlReturn           As IXMLDOMNode
    'BMIDS00190
    Dim xmlCVBHClassDefDoc  As FreeThreadedDOMDocument40
    Dim xmlCVBHList         As IXMLDOMNode
    Dim xmlBankruptcy       As IXMLDOMElement
    Dim xmlCVBH             As IXMLDOMElement
    Dim xmlRequest          As IXMLDOMElement
    Dim strGuid             As String
    Dim strNewGUID          As String
    Dim objSQLAssist        As SQLAssist
    Dim xmlTempNode         As IXMLDOMNode
    Dim xmlNode             As IXMLDOMNode
    'BMIDS00190 End
                  
    Set objICFClassDef = New CFClassDef
    Set xmlClassDefDoc = objICFClassDef.LoadBankruptcyHistoryData
    
    'BMIDS00190
    Set objSQLAssist = New SQLAssist
    Set xmlBankruptcy = xmlClassDefDoc.createElement("BANKRUPTCYHISTORYLIST")
    Set xmlCVBHClassDefDoc = objICFClassDef.LoadCustVerBankruptcyHistoryData()
    
    Set xmlCVBHList = m_objIDOAssist.FindListMultiple(vxmlTableElement, xmlCVBHClassDefDoc, "BANKRUPTCYHISTORYGUID, CUSTOMERNUMBER", "BANKRUPTCYHISTORY")
    strGuid = vbNullString
    
    For Each xmlCVBH In xmlCVBHList.selectNodes("CUSTOMERVERSIONBANKRUPTCYHISTORY")
        strNewGUID = m_objXmlAssist.GetMandatoryElementText(xmlCVBH, "BANKRUPTCYHISTORYGUID")
        If strGuid <> strNewGUID Then
            strGuid = strNewGUID
            ' Create new request node using the current GUID
            Set xmlRequest = xmlClassDefDoc.createElement("BANKRUPTCYHISTORYLIST")
            Set xmlTempNode = xmlClassDefDoc.createElement("BANKRUPTCYHISTORY")
            xmlRequest.appendChild xmlTempNode
            Set xmlTempNode = xmlClassDefDoc.createElement("BANKRUPTCYHISTORYGUID")
            xmlTempNode.Text = objSQLAssist.FormatGuid(strGuid)
            xmlRequest.firstChild.appendChild xmlTempNode
            ' Query the details of the Bankruptcy History record matching the current GUID
            Set xmlTempNode = m_objIDOAssist.FindListMultiple(xmlRequest, xmlClassDefDoc, vstrItemName:="BANKRUPTCYHISTORY")
            Set xmlNode = xmlTempNode.firstChild
            xmlBankruptcy.appendChild xmlNode
        End If
        xmlCVBH.removeChild xmlCVBH.selectSingleNode("BANKRUPTCYHISTORYGUID")
        xmlNode.appendChild xmlCVBH.cloneNode(True)
    Next
    
    Set xmlReturn = xmlBankruptcy
    'BMIDS00190 End
    
    'Set xmlReturn = m_objIDOAssist.FindListMultiple(vxmlTableElement, xmlClassDefDoc, _
    '        vstrOrderByField:="CUSTOMERNUMBER,CUSTOMERVERSIONNUMBER,SEQUENCENUMBER")

    OrderListByApplicant vxmlTableElement, xmlReturn, "BANKRUPTCYHISTORY"

    Set ICustomerFinancialDO_FindBankruptcyHistoryList = xmlReturn

    Set objICFClassDef = Nothing
    Set xmlClassDefDoc = Nothing
    Set xmlReturn = Nothing
    'BMIDS00190
    Set xmlCVBHClassDefDoc = Nothing
    Set xmlCVBHList = Nothing
    Set xmlBankruptcy = Nothing
    Set xmlCVBH = Nothing
    Set xmlRequest = Nothing
    Set objSQLAssist = Nothing
    Set xmlTempNode = Nothing
    Set xmlNode = Nothing
    'BMIDs00190 End
    
    If Not m_objContext Is Nothing Then
        m_objContext.SetComplete
    End If
    
    Exit Function

FindBankruptcyHistoryListVbErr:
    
    Set objICFClassDef = Nothing
    Set xmlClassDefDoc = Nothing
    Set xmlReturn = Nothing
    'BMIDS00190
    Set xmlCVBHClassDefDoc = Nothing
    Set xmlCVBHList = Nothing
    Set xmlBankruptcy = Nothing
    Set xmlCVBH = Nothing
    Set xmlRequest = Nothing
    Set objSQLAssist = Nothing
    Set xmlTempNode = Nothing
    Set xmlNode = Nothing
    'BMIDS00190 End
    
    m_objErrAssist.AddToErrSource strFunctionName
    
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

Private Function ICustomerFinancialDO_FindCCJHistoryList(ByVal vxmlTableElement As IXMLDOMElement) As IXMLDOMNode
' header ----------------------------------------------------------------------------------
' description:
'   Get the data for all instances of the persistant data associated with
'   this data object for the values supplied
' pass:
'   vxmlxmlTableElement  xml element containing the request
' return:                xml node containing retrieved data
'------------------------------------------------------------------------------------------
On Error GoTo FindCCJHistoryListVbErr

    Const strFunctionName As String = "ICustomerFinancialDO_FindCCJHistoryList"
    
    Dim xmlClassDefDoc          As FreeThreadedDOMDocument40
    Dim objICFClassDef          As ICFClassDef

    Dim intTotalCCJs            As Integer
    Dim lngTotalJudgementAmount As Long
    Dim xmlReturn               As IXMLDOMNode
    Dim xmlCCJNodeList          As IXMLDOMNodeList
    Dim xmlNode                 As IXMLDOMNode
    
    'BMIDS00190
    Dim xmlCVCHClassDefDoc      As FreeThreadedDOMDocument40
    Dim xmlCVCHList             As IXMLDOMNode
    Dim xmlCCJ                  As IXMLDOMElement
    Dim xmlCVCH                 As IXMLDOMElement
    Dim xmlRequest              As IXMLDOMElement
    Dim strGuid                 As String
    Dim strNewGUID              As String
    Dim objSQLAssist            As SQLAssist
    Dim xmlTempNode             As IXMLDOMNode
    'BMIDS00190 End
         
    Set objICFClassDef = New CFClassDef
    Set xmlClassDefDoc = objICFClassDef.LoadCCJHistoryData()
    
    'BMIDS00190
    Set objSQLAssist = New SQLAssist
    Set xmlCCJ = xmlClassDefDoc.createElement("CCJHISTORYLIST")
    Set xmlCVCHClassDefDoc = objICFClassDef.LoadCustVerCCJHistoryData()
    Set xmlCVCHList = m_objIDOAssist.FindListMultiple(vxmlTableElement, xmlCVCHClassDefDoc, "CCJHISTORYGUID, CUSTOMERNUMBER", "CCJHISTORY")
    strGuid = vbNullString
    
    For Each xmlCVCH In xmlCVCHList.selectNodes("CUSTOMERVERSIONCCJHISTORY")
        strNewGUID = m_objXmlAssist.GetMandatoryElementText(xmlCVCH, "CCJHISTORYGUID")
        If strGuid <> strNewGUID Then
            strGuid = strNewGUID
            ' Create new request node using the current GUID
            Set xmlRequest = xmlClassDefDoc.createElement("CCJHISTORYLIST")
            Set xmlTempNode = xmlClassDefDoc.createElement("CCJHISTORY")
            xmlRequest.appendChild xmlTempNode
            Set xmlTempNode = xmlClassDefDoc.createElement("CCJHISTORYGUID")
            xmlTempNode.Text = objSQLAssist.FormatGuid(strGuid)
            xmlRequest.firstChild.appendChild xmlTempNode
            ' Query the details of the CCJ History record matching the current GUID
            Set xmlTempNode = m_objIDOAssist.FindListMultiple(xmlRequest, xmlClassDefDoc, vstrItemName:="CCJHISTORY")
            Set xmlNode = xmlTempNode.firstChild
            xmlCCJ.appendChild xmlNode
        End If
        xmlCVCH.removeChild xmlCVCH.selectSingleNode("CCJHISTORYGUID")
        xmlNode.appendChild xmlCVCH.cloneNode(True)
    Next
    
    Set xmlReturn = xmlCCJ
    'BMIDS00190 End
    
    'Set xmlReturn = m_objIDOAssist.FindListMultiple(vxmlTableElement, xmlClassDefDoc, _
    '        vstrOrderByField:="CUSTOMERNUMBER,CUSTOMERVERSIONNUMBER,SEQUENCENUMBER")

    OrderListByApplicant vxmlTableElement, xmlReturn, "CCJHISTORY"

    ' Get derived totals
    Set xmlCCJNodeList = xmlReturn.selectNodes("CCJHISTORY")

    intTotalCCJs = xmlCCJNodeList.length
    For Each xmlNode In xmlCCJNodeList
        lngTotalJudgementAmount = lngTotalJudgementAmount + Val(xmlNode.selectSingleNode("VALUEOFJUDGEMENT").Text)
    Next xmlNode

    xmlReturn.appendChild(xmlReturn.ownerDocument.createElement("TOTALCCJS")).Text = CStr(intTotalCCJs)
    xmlReturn.appendChild(xmlReturn.ownerDocument.createElement("TOTALJUDGEMENTAMOUNT")).Text = Format$(CStr(lngTotalJudgementAmount), "0")

    Set ICustomerFinancialDO_FindCCJHistoryList = xmlReturn

    Set objICFClassDef = Nothing
    Set xmlClassDefDoc = Nothing
    Set xmlReturn = Nothing
    Set xmlNode = Nothing
    Set xmlCCJNodeList = Nothing
    'BMIDS00190
    Set xmlCVCHClassDefDoc = Nothing
    Set xmlCVCHList = Nothing
    Set xmlCCJ = Nothing
    Set xmlCVCH = Nothing
    Set xmlRequest = Nothing
    Set objSQLAssist = Nothing
    Set xmlTempNode = Nothing
    'BMIDS00190 End
    
    If Not m_objContext Is Nothing Then
        m_objContext.SetComplete
    End If
    
    Exit Function

FindCCJHistoryListVbErr:
    
    Set objICFClassDef = Nothing
    Set xmlClassDefDoc = Nothing
    Set xmlReturn = Nothing
    Set xmlNode = Nothing
    Set xmlCCJNodeList = Nothing
    'BMIDS00190
    Set xmlCVCHClassDefDoc = Nothing
    Set xmlCVCHList = Nothing
    Set xmlCCJ = Nothing
    Set xmlCVCH = Nothing
    Set xmlRequest = Nothing
    Set objSQLAssist = Nothing
    Set xmlTempNode = Nothing
    'BMIDS00190 End
    
    m_objErrAssist.AddToErrSource strFunctionName
    
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

Private Function ICustomerFinancialDO_FindDeclinedMortgageList(ByVal vxmlTableElement As IXMLDOMElement) As IXMLDOMNode
' header ----------------------------------------------------------------------------------
' description:
'   Get the data for all instances of the persistant data associated with
'   this data object for the values supplied
' pass:
'   vxmlxmlTableElement  xml element containing the request
' return:                xml node containing retrieved data
'------------------------------------------------------------------------------------------
On Error GoTo FindDeclinedMortgageListVbErr

    Const strFunctionName As String = "ICustomerFinancialDO_FindDeclinedMortgageList"

    Dim objErrAssist As New ErrAssist
    
    Dim objIDOAssist As IDOAssist
    Dim xmlClassDefDoc As FreeThreadedDOMDocument40
    Dim objICFClassDef As ICFClassDef
    Dim xmlReturn As IXMLDOMNode
    
    Dim objContext As ObjectContext
    Set objContext = GetObjectContext()
         
    If objContext Is Nothing Then
        Set objIDOAssist = New DOAssist
    Else
        Set objIDOAssist = objContext.CreateInstance(gstrBASE_COMPONENT & ".DOAssist")
    End If
     
    Set objICFClassDef = New CFClassDef
    Set xmlClassDefDoc = objICFClassDef.LoadDeclinedMortgageData()
    Set xmlReturn = objIDOAssist.FindListMultiple(vxmlTableElement, xmlClassDefDoc, _
            vstrOrderByField:="CUSTOMERNUMBER,CUSTOMERVERSIONNUMBER,SEQUENCENUMBER")

    OrderListByApplicant vxmlTableElement, xmlReturn, "DECLINEDMORTGAGE"

    Set ICustomerFinancialDO_FindDeclinedMortgageList = xmlReturn

    Set objICFClassDef = Nothing
    Set xmlClassDefDoc = Nothing
    Set objErrAssist = Nothing
    Set objIDOAssist = Nothing
    Set xmlReturn = Nothing
    
    If Not objContext Is Nothing Then
        objContext.SetComplete
    End If

    Exit Function

FindDeclinedMortgageListVbErr:
    
    Set objICFClassDef = Nothing
    Set xmlClassDefDoc = Nothing
    Set objIDOAssist = Nothing
    Set xmlReturn = Nothing
    
    objErrAssist.AddToErrSource strFunctionName
    
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

Private Function ICustomerFinancialDO_FindGroupConnectionList(ByVal vxmlTableElement As IXMLDOMElement) As IXMLDOMNode
' header ----------------------------------------------------------------------------------
' description:
'   Get the data for all instances of the persistant data associated with
'   this data object for the values supplied
' pass:
'   vxmlxmlTableElement  xml element containing the request
' return:                xml node containing retrieved data
'------------------------------------------------------------------------------------------
On Error GoTo FindGroupConnectionListVbErr

    Const strFunctionName As String = "ICustomerFinancialDO_FindGroupConnectionList"

    Dim objErrAssist As New ErrAssist
    
    Dim objIDOAssist As IDOAssist
    Dim xmlClassDefDoc As FreeThreadedDOMDocument40
    Dim objICFClassDef As ICFClassDef
    
    Dim objContext As ObjectContext
    Set objContext = GetObjectContext()
         
    If objContext Is Nothing Then
        Set objIDOAssist = New DOAssist
    Else
        Set objIDOAssist = objContext.CreateInstance(gstrBASE_COMPONENT & ".DOAssist")
    End If
     
    Set objICFClassDef = New CFClassDef
    Set xmlClassDefDoc = objICFClassDef.LoadGroupConnectionData
    Set ICustomerFinancialDO_FindGroupConnectionList = objIDOAssist.FindListMultiple(vxmlTableElement, xmlClassDefDoc, _
            vstrOrderByField:="GROUPCONNECTIONSEQUENCENUMBER")

    Set objICFClassDef = Nothing
    Set xmlClassDefDoc = Nothing
    Set objErrAssist = Nothing
    Set objIDOAssist = Nothing
    
    If Not objContext Is Nothing Then
        objContext.SetComplete
    End If
    
    Exit Function

FindGroupConnectionListVbErr:
    
    Set objICFClassDef = Nothing
    Set xmlClassDefDoc = Nothing
    Set objIDOAssist = Nothing
    
    objErrAssist.AddToErrSource strFunctionName
    
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

Private Function ICustomerFinancialDO_FindLiabilityList(ByVal vxmlTableElement As IXMLDOMElement) As IXMLDOMNode
' header ----------------------------------------------------------------------------------
' description:
'   Get the data for all instances of the persistant data associated with
'   this data object for the values supplied
' pass:
'   vxmlxmlTableElement  xml element containing the request
' return:                xml node containing retrieved data
'------------------------------------------------------------------------------------------
On Error GoTo FindLiabilityListVbErr

    Const strFunctionName As String = "ICustomerFinancialDO_FindLiabilityList"
    
    Dim xmlClassDefDoc As FreeThreadedDOMDocument40
    Dim objICFClassDef As ICFClassDef
    
    Dim xmlReturn As IXMLDOMNode
    Dim xmlTempNode As IXMLDOMNode
    Dim xmlNode As IXMLDOMNode

    Dim dblMonthlyRepayment        As Double, _
        dblTotalCurrentLiabilities As Double, _
        dblTotalFutureLiabilities  As Double, _
        dblSeqNum                  As Double
    
    'BMIDS00190
    Dim xmlCustomers            As IXMLDOMNode
    Dim xmlCustomer             As IXMLDOMNode
    Dim xmlLoan                 As IXMLDOMNode
    Dim xmlRequest              As IXMLDOMNode
    Dim xmlDirectory            As IXMLDOMNode
    Dim xmlThirdParty           As IXMLDOMNode
    Dim strAccountGUID          As String
    Dim xmlDirectoryClassDef    As FreeThreadedDOMDocument40
    Dim xmlThirdPartyClassDef   As FreeThreadedDOMDocument40
    'BMIDS00190 End
    
    Set objICFClassDef = New CFClassDef

    'BMIDS00190
    Set xmlClassDefDoc = objICFClassDef.LoadCFLoansLiabilitiesData()
    Set xmlDirectoryClassDef = objICFClassDef.LoadLoansLiabilitiesDirectoryData()
    Set xmlThirdPartyClassDef = objICFClassDef.LoadLoansLiabilitiesThirdPartyData()
    
    Set xmlReturn = xmlClassDefDoc.createElement("LOANSLIABILITIESLIST")
    Set xmlCustomers = m_objIDOAssist.FindListMultiple(vxmlTableElement, xmlClassDefDoc, "ACCOUNTGUID, CUSTOMERNUMBER", "LOANSLIABILITIES")
    
    If Not xmlCustomers Is Nothing Then      'BMIDS00383
    
        For Each xmlCustomer In xmlCustomers.selectNodes("CF_LOANSLIABILITIES")
            m_objXmlAssist.ChangeNodeName xmlCustomer, "CF_LOANSLIABILITIES", "LOANSLIABILITIES"
            strAccountGUID = xmlCustomer.selectSingleNode("ACCOUNTGUID").Text
            Set xmlLoan = xmlReturn.selectSingleNode("LOANSLIABILITIES[ACCOUNTGUID[.='" & strAccountGUID & "']]")
            If xmlLoan Is Nothing Then
                Set xmlLoan = xmlCustomer
                xmlReturn.appendChild xmlLoan
                
                Set xmlNode = xmlLoan.selectSingleNode("ACCOUNT/DIRECTORYGUID")
                If Len(Trim(xmlNode.Text)) > 0 Then
                    Set xmlRequest = xmlClassDefDoc.createElement("LOANSLIABILITIESDIRECTORY")
                    xmlRequest.appendChild xmlNode.cloneNode(True)
                    Set xmlDirectory = m_objIDOAssist.FindList(xmlRequest, xmlDirectoryClassDef)
                    xmlLoan.appendChild xmlDirectory
                End If
                
                Set xmlNode = xmlLoan.selectSingleNode("ACCOUNT/THIRDPARTYGUID")
                If Len(Trim(xmlNode.Text)) > 0 Then
                    Set xmlRequest = xmlClassDefDoc.createElement("LOANSLIABILITIESTHIRDPARTY")
                    xmlRequest.appendChild xmlNode.cloneNode(True)
                    Set xmlThirdParty = m_objIDOAssist.FindList(xmlRequest, xmlThirdPartyClassDef)
                    xmlLoan.appendChild xmlThirdParty
                End If
            Else
                xmlLoan.appendChild xmlCustomer.selectSingleNode("ACCOUNTRELATIONSHIP")
            End If
        Next
    'BMIDS00190 End
    
    End If  'BMIDS00383

'    ' Fetch LOANSLIABILITIESDIRECTORY records
'    Set xmlClassDefDoc = objICFClassDef.LoadLoansLiabilitiesDirectoryData
'    Set xmlTempNode = m_objIDOAssist.FindListMultiple(vxmlTableElement, xmlClassDefDoc, _
'            vstrOrderByField:="CUSTOMERNUMBER,CUSTOMERVERSIONNUMBER,SEQUENCENUMBER", vstrItemName:="LOANSLIABILITIES")
'    If Not xmlTempNode Is Nothing Then
'        Set xmlReturn = xmlTempNode
'    End If
'
'    ' Fetch LOANSLIABILITIESTHIRDPARTY records
'    Set xmlClassDefDoc = objICFClassDef.LoadLoansLiabilitiesThirdPartyData
'    Set xmlTempNode = Nothing
'    Set xmlTempNode = m_objIDOAssist.FindListMultiple(vxmlTableElement, xmlClassDefDoc, _
'            vstrOrderByField:="CUSTOMERNUMBER,CUSTOMERVERSIONNUMBER,SEQUENCENUMBER", vstrItemName:="LOANSLIABILITIES")
'
'    ' Add all the LLThirdParty records to the LLDirectory XML
'    If Not xmlTempNode Is Nothing Then
'        If Not xmlReturn Is Nothing Then
'            For Each xmlNode In xmlTempNode.childNodes
'                xmlReturn.appendChild xmlNode.cloneNode(True)
'            Next xmlNode
'        Else
'            Set xmlReturn = xmlTempNode
'        End If
'    End If
'
'    Set xmlTempNode = Nothing
'    Set xmlTempNode = ICustomerFinancialDO_FindLoansAndLiabilitiesWithNoLender(vxmlTableElement)
'
'    ' Add all the LOANSLIABILITIES records to the LLDirectory XML
'    If Not xmlTempNode Is Nothing Then
'        If Not xmlReturn Is Nothing Then
'            For Each xmlNode In xmlTempNode.childNodes
'                xmlReturn.appendChild xmlNode.cloneNode(True)
'            Next xmlNode
'        Else
'            Set xmlReturn = xmlTempNode
'        End If
'    End If
'
    If Not xmlReturn Is Nothing Then

        'JR - Omiplus24, append ContactTelephoneDetails Data to each ContactDetailsNode found
        Dim xmlCTNode As IXMLDOMNode
        Dim xmlCTNodeList As IXMLDOMNodeList
        Set xmlCTNodeList = xmlReturn.selectNodes(".//CONTACTDETAILS")
        For Each xmlCTNode In xmlCTNodeList
            Call AppendContactTelephoneData(xmlCTNode)
        Next
        'JR - End
'
'        For Each xmlNode In xmlReturn.selectNodes("LOANSLIABILITIESLIST/LOANSLIABILITIES")
'            m_objXmlAssist.PromoteNode xmlReturn, xmlNode
'        Next xmlNode
'        For Each xmlNode In xmlReturn.selectNodes("LOANSLIABILITIESDIRECTORY/LOANSLIABILITIES")
'            m_objXmlAssist.PromoteNode xmlReturn, xmlNode
'        Next xmlNode
'        For Each xmlNode In xmlReturn.selectNodes("LOANSLIABILITIESTHIRDPARTY/LOANSLIABILITIES")
'            m_objXmlAssist.PromoteNode xmlReturn, xmlNode
'        Next xmlNode
'

        OrderListByApplicant vxmlTableElement, xmlReturn, "LOANSLIABILITIES"

        '
        ' Calculate derived data
        '
        
        'dblSeqNum = 999    'BMIDS00190
        For Each xmlNode In xmlReturn.childNodes
            dblMonthlyRepayment = 0
    
            ' Get monthly repayment
            Set xmlTempNode = xmlNode.selectSingleNode(".//MONTHLYREPAYMENT")
            If Not (xmlTempNode Is Nothing) Then
                dblMonthlyRepayment = Val(xmlTempNode.Text)
            End If
    
            'BMIDS00190 Sequence number is no longer used
            'Set xmlTempNode = xmlNode.selectSingleNode(".//SEQUENCENUMBER")
            'If Not Val(xmlTempNode.Text) = dblSeqNum Then
            '    dblSeqNum = Val(xmlTempNode.Text)
            'BMIDS00190 End
            
                ' Add to current liabilities
                dblTotalCurrentLiabilities = dblTotalCurrentLiabilities + dblMonthlyRepayment
    
                ' Add to future liabilities
                Set xmlTempNode = xmlNode.selectSingleNode(".//LOANREPAYMENTINDICATOR")
                If Not (xmlTempNode Is Nothing) Then
                    If xmlTempNode.Text = "0" Then
                        dblTotalFutureLiabilities = dblTotalFutureLiabilities + dblMonthlyRepayment
                    End If
                End If
            'End If     'BMIDS00190
        Next xmlNode
    
        ' Append the totals to the XML
        xmlReturn.appendChild(xmlReturn.ownerDocument.createElement("TOTALCURRENTLIABILITIES")).Text = _
            Format(dblTotalCurrentLiabilities, "0.00")
    
        xmlReturn.appendChild(xmlReturn.ownerDocument.createElement("TOTALFUTURELIABILITIES")).Text = _
            Format(dblTotalFutureLiabilities, "0.00")

    End If

    Set ICustomerFinancialDO_FindLiabilityList = xmlReturn

    Set objICFClassDef = Nothing
    Set xmlClassDefDoc = Nothing
    Set xmlReturn = Nothing
    Set xmlNode = Nothing
    Set xmlTempNode = Nothing
    'JR - Omiplus24
    Set xmlCTNode = Nothing
    Set xmlCTNodeList = Nothing
    'BMIDS00190
    Set xmlCustomers = Nothing
    Set xmlCustomer = Nothing
    Set xmlLoan = Nothing
    Set xmlRequest = Nothing
    Set xmlDirectory = Nothing
    Set xmlThirdParty = Nothing
    Set xmlDirectoryClassDef = Nothing
    Set xmlThirdPartyClassDef = Nothing
    'BMIDS00190 End
    
    If Not m_objContext Is Nothing Then
        m_objContext.SetComplete
    End If
    
    Exit Function

FindLiabilityListVbErr:
    If m_objErrAssist.GetOmigaErrorNumber(Err.Number) = oeRecordNotFound Then
        Resume Next
    End If
    
    Set objICFClassDef = Nothing
    Set xmlClassDefDoc = Nothing
    Set xmlReturn = Nothing
    Set xmlNode = Nothing
    Set xmlTempNode = Nothing
    'JR - Omiplus24
    Set xmlCTNode = Nothing
    Set xmlCTNodeList = Nothing
    'BMIDS00190
    Set xmlCustomers = Nothing
    Set xmlCustomer = Nothing
    Set xmlLoan = Nothing
    Set xmlRequest = Nothing
    Set xmlDirectory = Nothing
    Set xmlThirdParty = Nothing
    Set xmlDirectoryClassDef = Nothing
    Set xmlThirdPartyClassDef = Nothing
    'BMIDS00190 End
    
    m_objErrAssist.AddToErrSource strFunctionName
    
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

Private Function ICustomerFinancialDO_FindLifeProductList(ByVal vxmlTableElement As IXMLDOMElement) As IXMLDOMNode
' header ----------------------------------------------------------------------------------
' description:
'   Get the data for all instances of the persistant data associated with
'   this data object for the values supplied
' pass:
'   vxmlxmlTableElement  xml element containing the request
' return:                xml node containing retrieved data
'------------------------------------------------------------------------------------------
On Error GoTo FindLifeProductListVbErr

    Const strFunctionName As String = "ICustomerFinancialDO_FindLifeProductList"

    Dim objErrAssist As New ErrAssist
    
    Dim objIDOAssist As IDOAssist, _
        objTPAssist As IThirdPartyAssist
    Set objTPAssist = New ThirdPartyAssist

    Dim xmlClassDefDoc As FreeThreadedDOMDocument40
    Dim objICFClassDef As ICFClassDef

    Dim xmlReturn As IXMLDOMNode
    Dim xmlNode As IXMLDOMNode
    Dim xmlNodeList As IXMLDOMNodeList

    Dim objContext As ObjectContext
    Set objContext = GetObjectContext()
         
    If objContext Is Nothing Then
        Set objIDOAssist = New DOAssist
    Else
        Set objIDOAssist = objContext.CreateInstance(gstrBASE_COMPONENT & ".DOAssist")
    End If
     
    Set objICFClassDef = New CFClassDef
    Set xmlClassDefDoc = objICFClassDef.LoadLifeProductData
    Set xmlReturn = objIDOAssist.FindListMultiple(vxmlTableElement, xmlClassDefDoc, _
            vstrOrderByField:="CUSTOMERNUMBER,CUSTOMERVERSIONNUMBER", vstrItemName:="LIFEPRODUCT")

    OrderListByApplicant vxmlTableElement, xmlReturn, "LIFEPRODUCT"

    Set xmlNodeList = xmlReturn.selectNodes(".//ACCOUNT")
    For Each xmlNode In xmlNodeList
        objTPAssist.GetThirdParty xmlNode
    Next xmlNode

    Set ICustomerFinancialDO_FindLifeProductList = xmlReturn

    Set objICFClassDef = Nothing
    Set xmlClassDefDoc = Nothing
    Set objErrAssist = Nothing
    Set objIDOAssist = Nothing
    Set xmlNode = Nothing
    Set xmlNodeList = Nothing
    Set xmlReturn = Nothing
    Set objTPAssist = Nothing
    
    If Not objContext Is Nothing Then
        objContext.SetComplete
    End If
    
    Exit Function

FindLifeProductListVbErr:
    
    Set objICFClassDef = Nothing
    Set xmlClassDefDoc = Nothing
    Set objIDOAssist = Nothing
    Set xmlNode = Nothing
    Set xmlNodeList = Nothing
    Set xmlReturn = Nothing
    Set objTPAssist = Nothing
    
    objErrAssist.AddToErrSource strFunctionName
    
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

Private Function ICustomerFinancialDO_FindMortgageAccountAndLoanList( _
                    ByVal vxmlTableElement As IXMLDOMElement) As IXMLDOMNode

On Error GoTo FindListVbErr

    Const strFunctionName As String = "ICustomerFinancialDO_FindMortgageAccountAndLoanList"

    Dim objErrAssist As ErrAssist
    Dim objXmlAssist As XMLAssist
    Set objErrAssist = New ErrAssist
    Set objXmlAssist = New XMLAssist
    
    Dim objICFClassDef As ICFClassDef
    Dim xmlClassDefDoc As FreeThreadedDOMDocument40
    Dim objIDOAssist As IDOAssist
    
    Dim xmlTempNode As IXMLDOMNode
    Dim xmlNode As IXMLDOMNode
    Dim xmlNodeList As IXMLDOMNodeList
        
'    Dim strCustomerNumber As String, strCustomerVersionNumber As String
'    Dim strCustomerAddressSequenceNumber As String, strSql As String, strCondition As String
        
    Dim objContext As ObjectContext
    Set objContext = GetObjectContext()
         
    If objContext Is Nothing Then
        Set objIDOAssist = New DOAssist
    Else
        Set objIDOAssist = objContext.CreateInstance(gstrBASE_COMPONENT & ".DOAssist")
    End If

    Set objICFClassDef = New CFClassDef
    Set xmlClassDefDoc = objICFClassDef.LoadMortgageAccountAndLoanData
    
    Set xmlTempNode = objIDOAssist.FindListMultiple(vxmlTableElement, xmlClassDefDoc, _
            vstrItemName:="MORTGAGEACCOUNTANDLOAN")

Set ICustomerFinancialDO_FindMortgageAccountAndLoanList = xmlTempNode

    Set objICFClassDef = Nothing
    Set xmlClassDefDoc = Nothing
    Set objErrAssist = Nothing
    Set objIDOAssist = Nothing
    Set xmlTempNode = Nothing
    Set xmlNode = Nothing
    Set objXmlAssist = Nothing
    
    If Not objContext Is Nothing Then
        objContext.SetComplete
    End If

    Exit Function
FindListVbErr:
    
    Set objICFClassDef = Nothing
    Set xmlClassDefDoc = Nothing
    Set objIDOAssist = Nothing
    Set xmlTempNode = Nothing
    Set xmlNode = Nothing
    Set objXmlAssist = Nothing
    
    objErrAssist.AddToErrSource strFunctionName
    
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

Private Function ICustomerFinancialDO_FindMortgageLoanList(ByVal vxmlTableElement As IXMLDOMElement) As IXMLDOMNode
' header ----------------------------------------------------------------------------------
' description:
'   Get the data for all instances of the persistant data associated with
'   this data object for the values supplied
' pass:
'   vxmlxmlTableElement  xml element containing the request
' return:                xml node containing retrieved data
'------------------------------------------------------------------------------------------
On Error GoTo FindMortgageLoanListVbErr

    Const strFunctionName As String = "ICustomerFinancialDO_FindMortgageLoanList"

    Dim objErrAssist As New ErrAssist
    Dim objXmlAssist As XMLAssist
    Set objXmlAssist = New XMLAssist
    
    Dim objIDOAssist As IDOAssist
    Dim xmlClassDefDoc As FreeThreadedDOMDocument40
    Dim objICFClassDef As ICFClassDef

    Dim xmlMortgageAccountNode As IXMLDOMNode, _
        xmlNode                As IXMLDOMNode
    Dim xmlReturn As IXMLDOMNode
    Dim objComboDO As New ComboDO

    Dim strRedemptionStatus As String
    Dim dblLoansNotRedeemed As Double, _
        dblPaymentsNotRedeemed As Double
    
    Dim objContext As ObjectContext
    Set objContext = GetObjectContext()
         
    If objContext Is Nothing Then
        Set objIDOAssist = New DOAssist
    Else
        Set objIDOAssist = objContext.CreateInstance(gstrBASE_COMPONENT & ".DOAssist")
    End If
     
    Set objICFClassDef = New CFClassDef
    Set xmlClassDefDoc = objICFClassDef.LoadMortgageLoanData
    
    ' SR 07/07/00 - SYS0892 : if OrderByStartDate is mentioned in Request tag, Order the
    '                         output in the Descending order of StartDate
    If objXmlAssist.GetTagValue(vxmlTableElement, "ORDERBYSTARTDATE") <> "1" Then
        Set xmlReturn = objIDOAssist.FindListMultiple(vxmlTableElement, xmlClassDefDoc)
    Else
        Set xmlReturn = objIDOAssist.FindListMultiple(vxmlTableElement, xmlClassDefDoc, " STARTDATE DESC")
    End If

    dblLoansNotRedeemed = 0
    dblPaymentsNotRedeemed = 0

    ' Calculate LOANSNOTREDEEMED and PAYMENTSNOTREDEEMED
    For Each xmlMortgageAccountNode In xmlReturn.childNodes
        strRedemptionStatus = ""

        Set xmlNode = xmlMortgageAccountNode.selectSingleNode(".//REDEMPTIONSTATUS")
        If Not (xmlNode Is Nothing) Then
            strRedemptionStatus = Trim$(xmlNode.Text)
        End If

        If strRedemptionStatus <> "" Then
            If objComboDO.IsItemInValidation("RedemptionStatus", strRedemptionStatus, "N") Or _
               objComboDO.IsItemInValidation("RedemptionStatus", strRedemptionStatus, "P") Then
                ' Outstanding balance
                Set xmlNode = xmlMortgageAccountNode.selectSingleNode(".//OUTSTANDINGBALANCE")
                If Not (xmlNode Is Nothing) Then
                    dblLoansNotRedeemed = dblLoansNotRedeemed + Val(xmlNode.Text)
                End If
    
                ' Monthly repayment
                Set xmlNode = xmlMortgageAccountNode.selectSingleNode(".//MONTHLYREPAYMENT")
                If Not (xmlNode Is Nothing) Then
                    dblPaymentsNotRedeemed = dblPaymentsNotRedeemed + Val(xmlNode.Text)
                End If
            End If
        End If
    Next xmlMortgageAccountNode

    'SYS0866 Standardise behaviour of Balance & Payment fields
    xmlReturn.appendChild(xmlReturn.ownerDocument.createElement("LOANSNOTREDEEMED")).Text = _
        Format$(CStr(dblLoansNotRedeemed), "0")
    xmlReturn.appendChild(xmlReturn.ownerDocument.createElement("PAYMENTSNOTREDEEMED")).Text = _
            Format$(CStr(dblPaymentsNotRedeemed), "0.00")

    Set ICustomerFinancialDO_FindMortgageLoanList = xmlReturn

    Set objICFClassDef = Nothing
    Set xmlClassDefDoc = Nothing
    Set objErrAssist = Nothing
    Set objXmlAssist = Nothing
    Set objIDOAssist = Nothing
    Set xmlMortgageAccountNode = Nothing
    Set xmlNode = Nothing
    Set xmlReturn = Nothing
    Set objComboDO = Nothing
    
    If Not objContext Is Nothing Then
        objContext.SetComplete
    End If
'
    Exit Function

FindMortgageLoanListVbErr:
    
    Set objICFClassDef = Nothing
    Set xmlClassDefDoc = Nothing
    Set objIDOAssist = Nothing
    Set xmlMortgageAccountNode = Nothing
    Set xmlNode = Nothing
    Set xmlReturn = Nothing
    Set objComboDO = Nothing
    Set objXmlAssist = Nothing
    
    objErrAssist.AddToErrSource strFunctionName
    
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



Private Function ICustomerFinancialDO_FindOutgoingsTotalList(ByVal vxmlTableElement As IXMLDOMElement) As IXMLDOMNode
' header ----------------------------------------------------------------------------------
' description:
'   Get the data for all instances of the persistant data associated with
'   this data object for the values supplied
' pass:
'   vxmlxmlTableElement  xml element containing the request
' return:                xml node containing retrieved data
'------------------------------------------------------------------------------------------
On Error GoTo FindOutgoingsTotalListVbErr

    Const strFunctionName As String = "ICustomerFinancialDO_FindOutgoingsTotalList"

    Dim objErrAssist As New ErrAssist
    
    Dim objIDOAssist As IDOAssist
    Dim xmlClassDefDoc As FreeThreadedDOMDocument40
    Dim objICFClassDef As ICFClassDef
    Dim xmlReturn As FreeThreadedDOMDocument40
    Set xmlReturn = New FreeThreadedDOMDocument40
    xmlReturn.validateOnParse = False
    xmlReturn.setProperty "NewParser", True
    Dim xmlNode As IXMLDOMNode
    Dim objXmlAssist As XMLAssist

    Dim xmlRequestNode As IXMLDOMNode
    Dim xmlResponseNode As IXMLDOMNode
    
    Dim objContext As ObjectContext
    Set objContext = GetObjectContext()
    Set objXmlAssist = New XMLAssist
         
    If objContext Is Nothing Then
        Set objIDOAssist = New DOAssist
    Else
        Set objIDOAssist = objContext.CreateInstance(gstrBASE_COMPONENT & ".DOAssist")
    End If
     
    Set objICFClassDef = New CFClassDef

    xmlReturn.appendChild xmlReturn.createElement("OUTGOINGSTOTALLIST")
    
    ' Retrieve BANKCREDITCARDLIST
    Set xmlResponseNode = Nothing
    Set xmlRequestNode = BuildRequestBlock(vxmlTableElement, "BANKCREDITCARD")
    Set xmlClassDefDoc = objICFClassDef.LoadBankCreditCardData
    Set xmlResponseNode = objIDOAssist.FindListMultiple(xmlRequestNode, xmlClassDefDoc, _
                vstrOrderByField:="CUSTOMERNUMBER,CUSTOMERVERSIONNUMBER,SEQUENCENUMBER")
    If Not (xmlResponseNode Is Nothing) Then
        xmlReturn.documentElement.appendChild xmlResponseNode
    End If

    ' Retrieve LOANSLIABILITIESLIST
    Set xmlResponseNode = Nothing
    Set xmlRequestNode = BuildRequestBlock(vxmlTableElement, "LOANSLIABILITIES")
    Set xmlClassDefDoc = objICFClassDef.LoadLoansLiabilitiesData
    Set xmlResponseNode = objIDOAssist.FindListMultiple(xmlRequestNode, xmlClassDefDoc, _
                vstrOrderByField:="CUSTOMERNUMBER,CUSTOMERVERSIONNUMBER,SEQUENCENUMBER")
    If Not (xmlResponseNode Is Nothing) Then
        xmlReturn.documentElement.appendChild xmlResponseNode
    End If

    ' Retrieve LIFEPRODUCTLIST
    Set xmlResponseNode = Nothing
    Set xmlRequestNode = BuildRequestBlock(vxmlTableElement, "LIFEPRODUCT")
    Set xmlClassDefDoc = objICFClassDef.LoadLifeProductData
    Set xmlResponseNode = objIDOAssist.FindListMultiple(xmlRequestNode, xmlClassDefDoc, _
            vstrOrderByField:="CUSTOMERNUMBER,CUSTOMERVERSIONNUMBER", vstrItemName:="LIFEPRODUCT")
    If Not (xmlResponseNode Is Nothing) Then
        xmlReturn.documentElement.appendChild xmlResponseNode
    End If

    ' Retrieve ARREARSHISTORYLIST
    Set xmlResponseNode = Nothing
    Set xmlRequestNode = BuildRequestBlock(vxmlTableElement, "ARREARSHISTORY")
    Set xmlClassDefDoc = objICFClassDef.LoadArrearsHistoryData
    Set xmlResponseNode = objIDOAssist.FindListMultiple(xmlRequestNode, xmlClassDefDoc, _
                vstrOrderByField:="CUSTOMERNUMBER,CUSTOMERVERSIONNUMBER,ARREARSSEQUENCENUMBER")
    If Not (xmlResponseNode Is Nothing) Then
        xmlReturn.documentElement.appendChild xmlResponseNode
    End If

    ' Retrieve BANKRUPTCYHISTORY
    Set xmlResponseNode = Nothing
    Set xmlRequestNode = BuildRequestBlock(vxmlTableElement, "BANKRUPTCYHISTORY")
    Set xmlClassDefDoc = objICFClassDef.LoadBankruptcyHistoryData
    Set xmlResponseNode = objIDOAssist.FindListMultiple(xmlRequestNode, xmlClassDefDoc, _
                vstrOrderByField:="CUSTOMERNUMBER,CUSTOMERVERSIONNUMBER,SEQUENCENUMBER")
    If Not (xmlResponseNode Is Nothing) Then
        xmlReturn.documentElement.appendChild xmlResponseNode
    End If

    ' Retrieve CCJHISTORY
    Set xmlResponseNode = Nothing
    Set xmlRequestNode = BuildRequestBlock(vxmlTableElement, "CCJHISTORY")
    Set xmlClassDefDoc = objICFClassDef.LoadCCJHistoryData
    Set xmlResponseNode = objIDOAssist.FindListMultiple(xmlRequestNode, xmlClassDefDoc, _
                vstrOrderByField:="CUSTOMERNUMBER,CUSTOMERVERSIONNUMBER,SEQUENCENUMBER")
    If Not (xmlResponseNode Is Nothing) Then
        xmlReturn.documentElement.appendChild xmlResponseNode
    End If

    Set ICustomerFinancialDO_FindOutgoingsTotalList = xmlReturn.documentElement

    Set objICFClassDef = Nothing
    Set xmlClassDefDoc = Nothing
    Set objErrAssist = Nothing
    Set objIDOAssist = Nothing
    Set xmlReturn = Nothing
    Set xmlNode = Nothing
    Set objXmlAssist = Nothing
    Set xmlRequestNode = Nothing
    Set xmlResponseNode = Nothing
    
    If Not objContext Is Nothing Then
        objContext.SetComplete
    End If
    
    Exit Function

FindOutgoingsTotalListVbErr:
    If objErrAssist.GetOmigaErrorNumber(Err.Number) = oeRecordNotFound Then
        Resume Next
    End If
    
    Set objICFClassDef = Nothing
    Set xmlClassDefDoc = Nothing
    Set objIDOAssist = Nothing
    Set xmlReturn = Nothing
    Set xmlNode = Nothing
    Set objXmlAssist = Nothing
    Set xmlRequestNode = Nothing
    Set xmlResponseNode = Nothing
    
    objErrAssist.AddToErrSource strFunctionName
    
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

Private Function ICustomerFinancialDO_FindRegularOutgoingsList(ByVal vxmlTableElement As IXMLDOMElement) As IXMLDOMNode
' header ----------------------------------------------------------------------------------
' description:
'   Get the data for all instances of the persistant data associated with
'   this data object for the values supplied
' pass:
'   vxmlxmlTableElement  xml element containing the request
' return:                xml node containing retrieved data
'------------------------------------------------------------------------------------------
On Error GoTo FindRegularOutgoingsListVbErr

    Const strFunctionName As String = "ICustomerFinancialDO_FindRegularOutgoingsList"
    
    Dim xmlClassDefDoc      As FreeThreadedDOMDocument40
    Dim objClassDef         As ICFClassDef
    Dim xmlNode             As IXMLDOMNode
    Dim xmlTempNode         As IXMLDOMNode
    
    'BMIDS00190
    
    Dim xmlCVROClassDefDoc  As FreeThreadedDOMDocument40
    Dim xmlCVROList         As IXMLDOMNode
    Dim xmlOutgoings        As IXMLDOMElement
    Dim xmlCVRO             As IXMLDOMElement
    Dim xmlRequest          As IXMLDOMElement
    Dim strGuid             As String
    Dim strNewGUID          As String
    Dim objSQLAssist        As SQLAssist
    Dim xmlDoc              As FreeThreadedDOMDocument40
    Dim xmlStylesheet       As FreeThreadedDOMDocument40
    Dim xmlTransformed      As FreeThreadedDOMDocument40
    'BMIDS00190 End
    
    Dim strXslPath As String
    Dim blnLoadedOk As Boolean

    Set objClassDef = New CFClassDef
    Set xmlClassDefDoc = objClassDef.LoadRegularOutgoingsData
    
    'BMIDS00190
    Set objSQLAssist = New SQLAssist
    Set xmlOutgoings = xmlClassDefDoc.createElement("REGULAROUTGOINGSLIST")
    Set xmlCVROClassDefDoc = objClassDef.LoadCustVerRegularOutgoingsData()
    Set xmlCVROList = m_objIDOAssist.FindListMultiple(vxmlTableElement, xmlCVROClassDefDoc, "REGULAROUTGOINGSGUID, CUSTOMERNUMBER", "REGULAROUTGOINGS")
    strGuid = vbNullString
    
    For Each xmlCVRO In xmlCVROList.selectNodes("CUSTOMERVERSIONREGULAROUTGOINGS")
        strNewGUID = m_objXmlAssist.GetMandatoryElementText(xmlCVRO, "REGULAROUTGOINGSGUID")
        If strGuid <> strNewGUID Then
            strGuid = strNewGUID
            ' Create new request node using the current GUID
            Set xmlRequest = xmlClassDefDoc.createElement("REGULAROUTGOINGSLIST")
            Set xmlTempNode = xmlClassDefDoc.createElement("REGULAROUTGOINGS")
            xmlRequest.appendChild xmlTempNode
            Set xmlTempNode = xmlClassDefDoc.createElement("REGULAROUTGOINGSGUID")
            xmlTempNode.Text = objSQLAssist.FormatGuid(strGuid)
            xmlRequest.firstChild.appendChild xmlTempNode
            ' Query the details of the Regular Outgoings record matching the current GUID
            Set xmlTempNode = m_objIDOAssist.FindListMultiple(xmlRequest, xmlClassDefDoc, vstrItemName:="REGULAROUTGOINGS")
            Set xmlNode = xmlTempNode.firstChild
            xmlOutgoings.appendChild xmlNode
        End If
        xmlCVRO.removeChild xmlCVRO.selectSingleNode("REGULAROUTGOINGSGUID")
        xmlNode.appendChild xmlCVRO.cloneNode(True)
    Next
    
    'Sort the result by RegularOutgoingsType
    Set xmlDoc = New FreeThreadedDOMDocument40
    xmlDoc.validateOnParse = False
    xmlDoc.setProperty "NewParser", True
    Set xmlTransformed = New FreeThreadedDOMDocument40
    xmlTransformed.validateOnParse = False
    xmlTransformed.setProperty "NewParser", True
    xmlDoc.appendChild xmlOutgoings
    
    Set xmlStylesheet = New FreeThreadedDOMDocument40
    xmlStylesheet.validateOnParse = False
    xmlStylesheet.setProperty "NewParser", True
    xmlStylesheet.async = False

    strXslPath = App.Path & "\FindRegularOutgoingsList.xslt"
    strXslPath = Replace(strXslPath, "DLL", "XML", 1, 1, vbTextCompare)
    blnLoadedOk = xmlStylesheet.load(strXslPath)

    If Not blnLoadedOk Then
        m_objErrAssist.ThrowError TypeName(Me), strFunctionName, oeSchemaNotLoaded, "Error in Loading XSLT File - " & strXslPath
    End If
    
    xmlDoc.transformNodeToObject xmlStylesheet.documentElement, xmlTransformed
    
    Set ICustomerFinancialDO_FindRegularOutgoingsList = xmlTransformed.documentElement
    'BMIDS00190 End
    
    'If (xmlReturn Is Nothing) Then
    '    objErrAssist.ThrowError TypeName(Me), strFunctionName, oeRecordNotfound
    'End If

    Set objClassDef = Nothing
    Set xmlClassDefDoc = Nothing
    Set xmlNode = Nothing
    Set xmlTempNode = Nothing
    'BMIDS00190
    Set xmlCVROClassDefDoc = Nothing
    Set xmlCVROList = Nothing
    Set xmlOutgoings = Nothing
    Set xmlCVRO = Nothing
    Set xmlRequest = Nothing
    Set objSQLAssist = Nothing
    'BMIDS00190 End
    
    If Not m_objContext Is Nothing Then
        m_objContext.SetComplete
    End If
    
    Exit Function

'RecordNotFoundErr:
'    If objErrAssist.GetOmigaErrorNumber(Err.Number) = oeRecordNotFound Then
'        Resume Next
'    End If

FindRegularOutgoingsListVbErr:
    
    Set objClassDef = Nothing
    Set xmlClassDefDoc = Nothing
    Set xmlNode = Nothing
    Set xmlTempNode = Nothing
    'BMIDS00190
    Set xmlCVROClassDefDoc = Nothing
    Set xmlCVROList = Nothing
    Set xmlOutgoings = Nothing
    Set xmlCVRO = Nothing
    Set xmlRequest = Nothing
    Set objSQLAssist = Nothing
    'BMIDS00190 End
    
    m_objErrAssist.AddToErrSource strFunctionName
    
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

Private Function ICustomerFinancialDO_GetLoansNotRedeemedTotals(ByVal vxmlTableElement As IXMLDOMElement) As IXMLDOMNode
' header ----------------------------------------------------------------------------------
' description:  Gets the total of outstanding balances and the total or repayments being
'               made on loans that are not to be redeemed' pass:
'   vxmlxmlTableElement  xml element containing the request
' return:                xml node containing retrieved data
'------------------------------------------------------------------------------------------
On Error GoTo GetLoansNotRedeemedTotalsVbErr

    Const strFunctionName As String = "ICustomerFinancialDO_GetLoansNotRedeemedTotals"

    Dim xmlCustomerNodeList As IXMLDOMNodeList
    Dim xmlCustomer As IXMLDOMElement
    Dim adoConnection As New ADODB.Connection
    Dim adoCommand As New ADODB.Command
    Dim rstThisRecSet As New ADODB.Recordset
    Dim adoParameter As ADODB.Parameter
    Dim objIADOAssist As IADOAssist
    Dim objIDOAssist As IDOAssist
    Dim objIomCFClassDef As ICFClassDef
    Dim xmlClassDefDoc As FreeThreadedDOMDocument40

    Dim strCustomerNo As String
    Dim strCustomerVersionNo As String

    If m_objContext Is Nothing Then
        Set objIADOAssist = New ADOAssist
        Set objIDOAssist = New DOAssist
    Else
        Set objIADOAssist = m_objContext.CreateInstance(gstrBASE_COMPONENT & ".ADOAssist")
        Set objIDOAssist = m_objContext.CreateInstance(gstrBASE_COMPONENT & ".DOAssist")
    End If

    Set xmlCustomerNodeList = vxmlTableElement.selectNodes(".//CUSTOMER")

    If xmlCustomerNodeList.length = 0 Then
        m_objErrAssist.ThrowError TypeName(Me), _
                                  strFunctionName, _
                                  oeMissingPrimaryTag, _
                                  "CUSTOMERVERSION tag missing"
    End If

    strCustomerNo = m_objXmlAssist.GetNodeValue(xmlCustomerNodeList.Item(0), _
                                                "./CUSTOMERNUMBER")
    strCustomerVersionNo = m_objXmlAssist.GetNodeValue(xmlCustomerNodeList.Item(0), _
                                                       "./CUSTOMERVERSIONNUMBER")

    Dim strSql As String

    ' PSC 22/07/02 BMIDS00006 - Start
    strSql = "SELECT SUM(MORTGAGELOAN.OUTSTANDINGBALANCE) LOANSNOTREDEEMED, " _
             & "SUM(MORTGAGELOAN.MONTHLYREPAYMENT) PAYMENTSONLOANSNOTREDEEMED " _
             & "FROM MORTGAGEACCOUNT, MORTGAGELOAN, ACCOUNTRELATIONSHIP AR, COMBOVALIDATION " _
             & "WHERE MORTGAGEACCOUNT.ACCOUNTGUID = MORTGAGELOAN.ACCOUNTGUID AND " _
             & "MORTGAGEACCOUNT.ACCOUNTGUID = AR.ACCOUNTGUID AND " _
             & "MORTGAGELOAN.REDEMPTIONSTATUS = COMBOVALIDATION.VALUEID AND " _
             & "COMBOVALIDATION.GROUPNAME = 'RedemptionStatus' AND " _
             & "COMBOVALIDATION.VALIDATIONTYPE = 'N' AND " _
             & "("

    adoCommand.ActiveConnection = objIADOAssist.GetConnStr

    Dim intIndex As Integer

    For intIndex = 0 To xmlCustomerNodeList.length - 1
        If intIndex > 0 Then
            strSql = strSql & " OR "
        End If

        strSql = strSql & "(AR.CUSTOMERNUMBER = ? AND AR.CUSTOMERVERSIONNUMBER = ?)"
        ' PSC 22/07/02 BMIDS00006 - End
        strCustomerNo = m_objXmlAssist.GetNodeValue(xmlCustomerNodeList.Item(intIndex), _
                                                    ".//CUSTOMERNUMBER")
        strCustomerVersionNo = m_objXmlAssist.GetNodeValue(xmlCustomerNodeList.Item(intIndex), _
                                                           ".//CUSTOMERVERSIONNUMBER")
        Set adoParameter = adoCommand.CreateParameter(, adBSTR, adParamInput, Len(strCustomerNo), strCustomerNo)
        adoCommand.Parameters.Append adoParameter

        Set adoParameter = adoCommand.CreateParameter(, adInteger, adParamInput)
        adoCommand.Parameters.Append adoParameter
        adoCommand((intIndex * 2) + 1) = CInt(strCustomerVersionNo)
    Next

    strSql = strSql & ")"

    adoCommand.CommandText = strSql

    Set rstThisRecSet = adoCommand.Execute(, , adCmdText)

    If rstThisRecSet.EOF Then
        ' raise application error to be interpreted by calling object
        m_objErrAssist.ThrowError TypeName(Me), strFunctionName, oeRecordNotFound
    End If

    Set objIomCFClassDef = New CFClassDef
    Set xmlClassDefDoc = objIomCFClassDef.LoadLoansNotRedeemedTotalsData()

    Set ICustomerFinancialDO_GetLoansNotRedeemedTotals = _
                            objIDOAssist.GetXMLFromRecordSet(rstThisRecSet, xmlClassDefDoc)

    rstThisRecSet.Close

    Set xmlCustomerNodeList = Nothing
    Set xmlCustomer = Nothing
    Set adoConnection = Nothing
    Set adoCommand = Nothing
    Set rstThisRecSet = Nothing
    Set adoParameter = Nothing
    Set objIADOAssist = Nothing
    Set objIDOAssist = Nothing
    Set objIomCFClassDef = Nothing
    Set xmlClassDefDoc = Nothing

    If Not m_objContext Is Nothing Then
        m_objContext.SetComplete
    End If

    Exit Function

GetLoansNotRedeemedTotalsVbErr:

    ' Close the recordset if still open
        If Not rstThisRecSet Is Nothing Then
                If rstThisRecSet.State = adStateOpen Then
                        rstThisRecSet.Close
                End If
        End If

    Set xmlCustomerNodeList = Nothing
    Set xmlCustomer = Nothing
    Set adoConnection = Nothing
    Set adoCommand = Nothing
    Set rstThisRecSet = Nothing
    Set adoParameter = Nothing
    Set objIADOAssist = Nothing
    Set objIDOAssist = Nothing
    Set objIomCFClassDef = Nothing
    Set xmlClassDefDoc = Nothing

    m_objErrAssist.AddToErrSource strFunctionName

    If m_objErrAssist.IsSystemError = True Then
        m_objErrAssist.LogError TypeName(Me), strFunctionName, Err.Number, Err.Description
    End If

    If Not m_objContext Is Nothing Then
        m_objContext.SetComplete
    End If

    '   re-raise error for business object to interpret as appropriate
    Err.Raise Err.Number, Err.Source, Err.Description

End Function

Private Function ICustomerFinancialDO_GetMtgAcctDataForFurtherAdv(ByVal vxmlRequest As IXMLDOMElement) As IXMLDOMNode
    
    On Error GoTo GetMtgAcctDataForFurtherAdvVbErr

    Const cstrFunctionName As String = "ICustomerFinancialDO_GetMtgAcctDataForFurtherAdv"
    
    Dim objIADOAssist               As IADOAssist
    Dim objConvAssist               As ConvertAssist
    Dim xmlResponseData             As IXMLDOMNode
    Dim xmlResponseElem             As IXMLDOMNode
    Dim xmlOut                      As FreeThreadedDOMDocument40
    Dim strSql                      As String
    Dim strApplicationNumber        As String
    'Dim strApplicationFactFindNumber As String
    'Dim strBMOnly                   As String
    Dim adoConnection               As ADODB.Connection
    Dim adoCommand                  As ADODB.Command
    Dim adoAppNumberParameter       As ADODB.Parameter
    'Dim adoAppFFNumberParameter     As ADODB.Parameter
    'Dim adoBMOnlyParameter          As ADODB.Parameter
    Dim adoTotalMonthlyRepayment    As ADODB.Parameter
    Dim adoTotalOutstandingBalance  As ADODB.Parameter
    Dim adoTotalIndemnityAmount     As ADODB.Parameter
    Dim dblMonthlyRepayment         As Double
    Dim dblOutstandingBalance       As Double
    Dim dblIndemnityAmount          As Double
    'EP2_9 GHun
    Dim adoLastValuationDate        As ADODB.Parameter
    Dim strLastValuationDate        As String
    'EP2_9 End
                       
    If m_objContext Is Nothing Then
        Set objIADOAssist = New ADOAssist
    Else
        Set objIADOAssist = m_objContext.CreateInstance(gstrBASE_COMPONENT & ".ADOAssist")
    End If
        
    ' Set MonthlyRepayment and Outstanding Balance from all MortgageLoan Records
    strApplicationNumber = Trim$(m_objXmlAssist.GetTagValue(vxmlRequest, "APPLICATIONNUMBER"))
    'strApplicationFactFindNumber = Trim$(m_objXmlAssist.GetTagValue(vxmlRequest, "APPLICATIONFACTFINDNUMBER"))
    
    'EP2_9 GHun redundant
    ''BMIDS591 GHun 02/07/2003 Default to false (0) if GetBMOnlyNotToBeRedeemedLoans not specified
    'strBMOnly = Trim$(m_objXmlAssist.GetTagValue(vxmlRequest, "GETBMONLYNOTTOBEREDEEMEDLOANS"))
    'If Len(strBMOnly) = 0 Then
    '    strBMOnly = "0"
    'End If
    
    'If Len(strApplicationNumber) = 0 Or Len(strApplicationFactFindNumber) = 0 Then
        'm_objErrAssist.ThrowError TypeName(Me), cstrFunctionName, oeInvalidParameter, _
        '                "ApplicationNumber & ApplicationFactFindNumber must be specified"
    If Len(strApplicationNumber) = 0 Then
        m_objErrAssist.ThrowError TypeName(Me), cstrFunctionName, oeInvalidParameter, _
                        "ApplicationNumber must be specified"
    End If
    'EP2_9 End
    
    #If GENERIC_SQL Then
        Dim eDbEngineType As DBENGINETYPE
        eDbEngineType = objIADOAssist.GetDBEngine
        If eDbEngineType = SQLServer Then
            strSql = "USP_GETMORTGAGELOANTOTALS"
        End If
    #End If
    
    Set adoConnection = New ADODB.Connection
    With adoConnection
        .ConnectionString = objIADOAssist.GetConnStr
        #If Not GENERIC_SQL Then
            .CursorLocation = adUseClient
        #End If
        .open
    End With
    
    Set adoCommand = New ADODB.Command
    
    With adoCommand
        Set .ActiveConnection = adoConnection
        .CommandType = adCmdStoredProc
        .CommandText = strSql
        
        'Application Number
        Set adoAppNumberParameter = .CreateParameter("APPLICATIONNUMBER", adBSTR, adParamInput, Len(strApplicationNumber), strApplicationNumber)
    
        'EP2_9 GHun redundant
        'Application Fact Find Number
        'Set adoAppFFNumberParameter = .CreateParameter("APPLICATIONFACTFINDNUMBER", adInteger, adParamInput, , CLng(strApplicationFactFindNumber))
       
        ''Get BMids only not to be redeemed loans
        'Set adoBMOnlyParameter = .CreateParameter("GETBMONLYNOTTOBEREDEEMEDLOANS", adChar, adParamInput, 1, strBMOnly)
        'EP2_9 GHun end
    
        'Output parameters
        Set adoTotalMonthlyRepayment = .CreateParameter("TOTALMONTHLYREPAYMENT", adDouble, adParamOutput)
        Set adoTotalOutstandingBalance = .CreateParameter("TOTALOUTSTANDINGBALANCE", adDouble, adParamOutput)
        Set adoTotalIndemnityAmount = .CreateParameter("TOTALINDEMNITYAMOUNT", adDouble, adParamOutput)
        Set adoLastValuationDate = .CreateParameter("LASTVALUATIONDATE", adDBTimeStamp, adParamOutput)  'EP2_9 GHun
    
        With .Parameters
            .Append adoAppNumberParameter
            '.Append adoAppFFNumberParameter    'EP2_9 GHun redundant
            '.Append adoBMOnlyParameter 'EP2_9 GHun redundant
            .Append adoTotalMonthlyRepayment
            .Append adoTotalOutstandingBalance
            .Append adoTotalIndemnityAmount
            .Append adoLastValuationDate    'EP2_9 GHun
        End With
        
        .Execute
        Set .ActiveConnection = Nothing
    End With
 
    adoConnection.Close
 
    Set objConvAssist = New ConvertAssist
    dblMonthlyRepayment = objConvAssist.CSafeDbl(adoTotalMonthlyRepayment.Value)
    dblOutstandingBalance = objConvAssist.CSafeDbl(adoTotalOutstandingBalance.Value)
    dblIndemnityAmount = objConvAssist.CSafeDbl(adoTotalIndemnityAmount.Value)
    'EP2_9 GHun
    If Not IsNull(adoLastValuationDate.Value) Then
        strLastValuationDate = objConvAssist.CSafeDate(adoLastValuationDate.Value)
    End If
    'EP2_9 End
    
    Set xmlOut = New FreeThreadedDOMDocument40
    xmlOut.validateOnParse = False
    xmlOut.setProperty "NewParser", True
    Set xmlResponseData = xmlOut.createElement("MORTGAGEACCOUNT")
       
    'Append to the response
    Set xmlResponseElem = xmlOut.createElement("TOTALMONTHLYREPAYMENT")
    xmlResponseElem.Text = dblMonthlyRepayment
    xmlResponseData.appendChild xmlResponseElem
    
    'Append to the response
    Set xmlResponseElem = xmlOut.createElement("TOTALOUTSTANDINGBALANCE")
    xmlResponseElem.Text = dblOutstandingBalance
    xmlResponseData.appendChild xmlResponseElem
    
    'Append to the response
    Set xmlResponseElem = xmlOut.createElement("TOTALORIGINALMIGAMOUNT")
    xmlResponseElem.Text = dblIndemnityAmount
    xmlResponseData.appendChild xmlResponseElem
    
    'EP2_9 GHun Append LastValuationDate to the response
    If Len(strLastValuationDate) > 0 Then
        Set xmlResponseElem = xmlOut.createElement("LASTVALUATIONDATE")
        xmlResponseElem.Text = strLastValuationDate
        xmlResponseData.appendChild xmlResponseElem
    End If
    'EP2_9 End
    
    Set ICustomerFinancialDO_GetMtgAcctDataForFurtherAdv = xmlResponseData
       
    m_objContext.SetComplete
    
    Set objIADOAssist = Nothing
    Set objConvAssist = Nothing
    Set xmlResponseData = Nothing
    Set xmlResponseElem = Nothing
    Set xmlOut = Nothing
    Set adoConnection = Nothing
    Set adoCommand = Nothing
    Set adoAppNumberParameter = Nothing
    'Set adoAppFFNumberParameter = Nothing
    'Set adoBMOnlyParameter = Nothing
    Set adoTotalMonthlyRepayment = Nothing
    Set adoTotalOutstandingBalance = Nothing
    Set adoTotalIndemnityAmount = Nothing
    Set adoLastValuationDate = Nothing
    
    Exit Function

GetMtgAcctDataForFurtherAdvVbErr:
    
    Set objIADOAssist = Nothing
    Set objConvAssist = Nothing
    Set xmlResponseData = Nothing
    Set xmlResponseElem = Nothing
    Set xmlOut = Nothing
    Set adoConnection = Nothing
    Set adoCommand = Nothing
    Set adoAppNumberParameter = Nothing
    'Set adoAppFFNumberParameter = Nothing
    'Set adoBMOnlyParameter = Nothing
    Set adoTotalMonthlyRepayment = Nothing
    Set adoTotalOutstandingBalance = Nothing
    Set adoTotalIndemnityAmount = Nothing
    Set adoLastValuationDate = Nothing
    
    m_objErrAssist.AddToErrSource (cstrFunctionName)
    
    If m_objErrAssist.IsSystemError = True Then
        m_objErrAssist.LogError TypeName(Me), cstrFunctionName, Err.Number, Err.Description
    End If
    
    If m_objErrAssist.IsSystemError Then
        m_objContext.SetAbort
    Else
        m_objContext.SetComplete
    End If
        
    '   re-raise error for business object to interpret as appropriate
    Err.Raise Err.Number, Err.Source, Err.Description
    
End Function

Private Function ICustomerFinancialDO_GetPortedStepAndPeriodFromMortgageAccount( _
                                    ByVal vxmlRequest As IXMLDOMElement) As IXMLDOMNode

On Error GoTo GetPortedStepAndPeriodFromMortgageAccountVbErr

    Const strFunctionName As String = "ICustomerFinancialDO_GetPortedStepAndPeriodFromMortgageAccount"
    
    Dim objIADOAssist  As IADOAssist
    Dim adoConnection As ADODB.Connection
    Dim adoCommand As ADODB.Command
    Dim adoStream As ADODB.Stream
    Dim xmlOut As FreeThreadedDOMDocument40
    Dim strApplicationNumber As String
    Dim strMSQNumber As String
    Dim strIsTOE As String  'MAR46 GHun
    Dim strMortgageGUID As String  'EP2_56 AShaw
    Dim objSQLAssist As SQLAssist 'EP2_56 AShaw

    Set objSQLAssist = New SQLAssist 'EP2_56 AShaw
    
    strApplicationNumber = m_objXmlAssist.GetTagValue(vxmlRequest, "APPLICATIONNUMBER")
    strMSQNumber = m_objXmlAssist.GetTagValue(vxmlRequest, "MORTGAGESUBQUOTENUMBER")
    strIsTOE = m_objXmlAssist.GetTagValue(vxmlRequest, "TOEIND")   'MAR46 GHun
    strMortgageGUID = m_objXmlAssist.GetTagValue(vxmlRequest, "MORTGAGELOANGUID")   'EP2_56 AShaw
    
    Set objIADOAssist = m_objContext.CreateInstance(gstrBASE_COMPONENT & ".ADOAssist")
    
    Set adoConnection = New ADODB.Connection
    With adoConnection
        .ConnectionString = objIADOAssist.GetConnStr
        .CursorLocation = adUseClient
        .open
    End With
    
    'BMIDS815 GHun The stored proc now returns XML for multiple LoanComponents,
    'rather than output parameters for just one LoanComponent
    
    Set adoStream = New ADODB.Stream
    adoStream.open

    Set adoCommand = New ADODB.Command
    With adoCommand
        .CommandText = "USP_GETPORTEDSTEPANDPERIODFROMACC"
        .CommandType = adCmdStoredProc
        .Parameters.Append .CreateParameter("@P_APPLICATIONNUMBER", adBSTR, adParamInput, Len(strApplicationNumber), strApplicationNumber)
        .Parameters.Append .CreateParameter("@P_MORTGAGESUBQUOTENUMBER", adInteger, adParamInput, , CInt(strMSQNumber))
        .Parameters.Append .CreateParameter("@P_ISTRANSFEROFEQUITY", adBoolean, adParamInput, , CInt(strIsTOE)) 'MAR46 GHun
        ' AShaw EP2_56
        If strMortgageGUID <> "" Then
            .Parameters.Append .CreateParameter("@p_MortgageGUID", adBinary, adParamInput, 16, objSQLAssist.GuidStringToByteArray(strMortgageGUID))
        End If
        Set .ActiveConnection = adoConnection
        .Properties("Output Stream") = adoStream
        .Properties("XML Root") = "RESPONSE"
        .Execute , , adExecuteStream
        Set .ActiveConnection = Nothing
    End With
    
    Set xmlOut = New FreeThreadedDOMDocument40
    xmlOut.validateOnParse = False
    xmlOut.setProperty "NewParser", True
    
    If xmlOut.loadXML(adoStream.ReadText) Then
        'ApplicationQuoteTxBO.GetPortedStepAndPeriodFromMortgageAccount handles the case of
        'record not found.
        'If xmlOut.selectSingleNode("RESPONSE/MORTGAGELOAN") Is Nothing Then
        '    m_objErrAssist.ThrowError TypeName(Me), strFunctionName, oeRecordNotFound
        'End If
    Else
        ' raise application error to be interpreted by calling object
        m_objErrAssist.ThrowError TypeName(Me), strFunctionName, oeRecordNotFound
    End If
    'BMIDS815 End
    
    adoConnection.Close
         
    Set ICustomerFinancialDO_GetPortedStepAndPeriodFromMortgageAccount = xmlOut.documentElement
    
    Set adoConnection = Nothing
    Set objIADOAssist = Nothing
    Set adoCommand = Nothing
    Set adoStream = Nothing
    Set xmlOut = Nothing
    
    Exit Function
    
GetPortedStepAndPeriodFromMortgageAccountVbErr:
    
    Set objIADOAssist = Nothing
    Set adoConnection = Nothing
    Set adoCommand = Nothing
    Set adoStream = Nothing
    Set xmlOut = Nothing
            
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

Private Sub ICustomerFinancialDO_UpdateArrearsHistory(ByVal vxmlTableElement As IXMLDOMElement)
' header ----------------------------------------------------------------------------------
' description:
'   UpdateArrearsHistory a single instance of the persistant data associated with this
'   data object
' pass:
'   vxmlxmlTableElement  xml element containing the request
' return:                n/a
'------------------------------------------------------------------------------------------
On Error GoTo UpdateArrearsHistoryVbErr

    Const strFunctionName As String = "ICustomerFinancialDO_UpdateArrearsHistory"
        
    Dim xmlClassDefDoc          As FreeThreadedDOMDocument40
    Dim objICFClassDef          As ICFClassDef
    'BMIDS00190
    Dim xmlAccountRelClassDef   As FreeThreadedDOMDocument40
    Dim xmlAccountClassDef      As FreeThreadedDOMDocument40
    Dim xmlOtherAAClassDef      As FreeThreadedDOMDocument40
    Dim xmlOtherArrears         As IXMLDOMNode
    Dim xmlCustomer             As IXMLDOMNode
    Dim xmlAccountGuid          As IXMLDOMNode
    Dim xmlRequest              As IXMLDOMNode
    Dim xmlDelete               As IXMLDOMNode
    Dim blnDeleteOldAccount     As Boolean
    Dim blnAccountGUIDExists    As Boolean
    'BMIDS00190 End
    
    Set objICFClassDef = New CFClassDef
    Set xmlClassDefDoc = objICFClassDef.LoadArrearsHistoryData()
    
    'BMIDS00190
    blnDeleteOldAccount = False
    
    Set xmlAccountRelClassDef = objICFClassDef.LoadAccountRelationshipData()
    Set xmlOtherAAClassDef = objICFClassDef.LoadOtherArrearsAccountData()
    Set xmlAccountClassDef = objICFClassDef.LoadAccountData()
    
    Set xmlOtherArrears = vxmlTableElement.selectSingleNode("OTHERARREARSACCOUNT")
    
    'Check if there is an existing Account GUID or not
    Set xmlAccountGuid = vxmlTableElement.selectSingleNode("ACCOUNTGUID")
    If xmlAccountGuid Is Nothing Then
        blnAccountGUIDExists = False
    ElseIf Len(Trim(xmlAccountGuid.Text)) = 0 Then
        blnAccountGUIDExists = False
        'Remove blank Account GUID node
        vxmlTableElement.removeChild xmlAccountGuid
    Else
        blnAccountGUIDExists = True
    End If
        
    'If there is no existing AccountGUID then a new one is required
    If Not blnAccountGUIDExists Then
        Set xmlAccountGuid = xmlClassDefDoc.createElement("ACCOUNTGUID")
        xmlAccountGuid.Text = CreateGUID()
        vxmlTableElement.appendChild xmlAccountGuid.cloneNode(True)
        
        'Create new Account
        Set xmlRequest = xmlClassDefDoc.createElement("ACCOUNT")
        xmlRequest.appendChild xmlAccountGuid.cloneNode(True)
        m_objIDOAssist.Create xmlRequest, xmlAccountClassDef
        
        'Create new Other Arrears Account
        xmlOtherArrears.appendChild xmlAccountGuid.cloneNode(True)
        m_objIDOAssist.Create xmlOtherArrears, xmlOtherAAClassDef
    Else
        If Not xmlOtherArrears Is Nothing Then
            'Update Other Arrears Account
            xmlOtherArrears.appendChild xmlAccountGuid.cloneNode(True)
            m_objIDOAssist.Update xmlOtherArrears, xmlOtherAAClassDef
        End If
    End If
    
    If Not xmlOtherArrears Is Nothing Then
        vxmlTableElement.removeChild xmlOtherArrears
    End If
    
    'Add new AccountRelationship links
    For Each xmlCustomer In vxmlTableElement.selectNodes("ACCOUNTRELATIONSHIP")
        xmlCustomer.appendChild xmlAccountGuid.cloneNode(True)
        m_objIDOAssist.Create xmlCustomer, xmlAccountRelClassDef
        vxmlTableElement.removeChild xmlCustomer
    Next
    
    Set xmlDelete = vxmlTableElement.selectSingleNode("DELETE")
    If Not xmlDelete Is Nothing Then
        Set xmlOtherArrears = xmlDelete.selectSingleNode("OTHERARREARSACCOUNT")
        If Not xmlOtherArrears Is Nothing Then
            'Remove old Other Arrears Account
            m_objIDOAssist.Delete xmlOtherArrears, xmlOtherAAClassDef
        
            'Remove AccountRelationship links from old Other Arrears Account
            Set xmlRequest = xmlClassDefDoc.createElement("ACCOUNTRELATIONSHIP")
            xmlRequest.appendChild xmlOtherArrears.selectSingleNode("ACCOUNTGUID").cloneNode(True)
            m_objIDOAssist.DeleteAll xmlRequest, xmlAccountRelClassDef
            
            'Old account will be deleted after update of arrears history record
            blnDeleteOldAccount = True
            Set xmlRequest = xmlClassDefDoc.createElement("ACCOUNT")
            xmlRequest.appendChild xmlOtherArrears.selectSingleNode("ACCOUNTGUID").cloneNode(True)
        End If
        
        'Remove old AccountRelationship links from current Account
        For Each xmlCustomer In xmlDelete.selectNodes("ACCOUNTRELATIONSHIP")
            xmlCustomer.appendChild xmlAccountGuid.cloneNode(True)
            m_objIDOAssist.Delete xmlCustomer, xmlAccountRelClassDef
        Next
        vxmlTableElement.removeChild xmlDelete
    End If
    'BMIDS00190 End
    
    m_objIDOAssist.Update vxmlTableElement, xmlClassDefDoc
    
    'BMIDS00190
    If blnDeleteOldAccount Then
        'The old account can only be deleted after ArrearsHistory has been updated
        'as the accountGUID is a foreign key
        m_objIDOAssist.Delete xmlRequest, xmlAccountClassDef
    End If
    'BMIDS00190 End
    
    Set objICFClassDef = Nothing
    Set xmlClassDefDoc = Nothing
    'BMIDS00190
    Set xmlAccountRelClassDef = Nothing
    Set xmlAccountClassDef = Nothing
    Set xmlOtherAAClassDef = Nothing
    Set xmlOtherArrears = Nothing
    Set xmlCustomer = Nothing
    Set xmlAccountGuid = Nothing
    Set xmlRequest = Nothing
    Set xmlDelete = Nothing
    'BMIDS00190 End
    
    If Not m_objContext Is Nothing Then
        m_objContext.SetComplete
    End If
    
    Exit Sub

UpdateArrearsHistoryVbErr:
        
    Set objICFClassDef = Nothing
    Set xmlClassDefDoc = Nothing
    'BMIDS00190
    Set xmlAccountRelClassDef = Nothing
    Set xmlAccountClassDef = Nothing
    Set xmlOtherAAClassDef = Nothing
    Set xmlOtherArrears = Nothing
    Set xmlCustomer = Nothing
    Set xmlAccountGuid = Nothing
    Set xmlRequest = Nothing
    Set xmlDelete = Nothing
    'BMIDS00190 End
    
    m_objErrAssist.AddToErrSource strFunctionName
    
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

Private Sub ICustomerFinancialDO_UpdateBankCard(ByVal vxmlTableElement As IXMLDOMElement)
' header ----------------------------------------------------------------------------------
' description:
'   UpdateBankCard a single instance of the persistant data associated with this
'   data object
' pass:
'   vxmlxmlTableElement  xml element containing the request
' return:                n/a
'------------------------------------------------------------------------------------------
On Error GoTo UpdateBankCardVbErr

    Const strFunctionName As String = "ICustomerFinancialDO_UpdateBankCard"
        
    Dim objErrAssist As New ErrAssist
    
    Dim objContext As ObjectContext
    Set objContext = GetObjectContext()
    
    Dim xmlClassDefDoc As FreeThreadedDOMDocument40
    Dim objICFClassDef As ICFClassDef
    Dim objIDOAssist As IDOAssist

    If objContext Is Nothing Then
       Set objIDOAssist = New DOAssist
    Else
       Set objIDOAssist = objContext.CreateInstance(gstrBASE_COMPONENT & ".DOAssist")
    End If
    
    Set objICFClassDef = New CFClassDef
    Set xmlClassDefDoc = objICFClassDef.LoadBankCreditCardData
    objIDOAssist.Update vxmlTableElement, xmlClassDefDoc
    
    Set objICFClassDef = Nothing
    Set xmlClassDefDoc = Nothing
    Set objIDOAssist = Nothing
    Set objErrAssist = Nothing
    
    If Not objContext Is Nothing Then
        objContext.SetComplete
    End If
    
    Exit Sub

UpdateBankCardVbErr:
        
    Set objICFClassDef = Nothing
    Set xmlClassDefDoc = Nothing
    Set objIDOAssist = Nothing
    
    objErrAssist.AddToErrSource strFunctionName
    
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

Private Sub ICustomerFinancialDO_UpdateBankruptcyHistory(ByVal vxmlTableElement As IXMLDOMElement)
' header ----------------------------------------------------------------------------------
' description:
'   UpdateBankruptcyHistory a single instance of the persistant data associated with this
'   data object
' pass:
'   vxmlxmlTableElement  xml element containing the request
' return:                n/a
'------------------------------------------------------------------------------------------
On Error GoTo UpdateBankruptcyHistoryVbErr

    Const strFunctionName As String = "ICustomerFinancialDO_UpdateBankruptcyHistory"
        
    Dim xmlClassDefDoc      As FreeThreadedDOMDocument40
    Dim objICFClassDef      As ICFClassDef
    'BMIDS00190
    Dim xmlCVBHClassDefDoc  As FreeThreadedDOMDocument40
    Dim xmlCustomer         As IXMLDOMElement
    Dim xmlDelete           As IXMLDOMNode
    'BMIDS00190 End
   
    Set objICFClassDef = New CFClassDef
    Set xmlClassDefDoc = objICFClassDef.LoadBankruptcyHistoryData()
    
    'BMIDS00190
    Set xmlCVBHClassDefDoc = objICFClassDef.LoadCustVerBankruptcyHistoryData()
    
    'Add new customer version links
    For Each xmlCustomer In vxmlTableElement.selectNodes("CUSTOMERVERSIONBANKRUPTCYHISTORY")
        xmlCustomer.appendChild vxmlTableElement.selectSingleNode("BANKRUPTCYHISTORYGUID").cloneNode(True)
        m_objIDOAssist.Create xmlCustomer, xmlCVBHClassDefDoc
        vxmlTableElement.removeChild xmlCustomer
    Next
    
    'Remove old customer version links
    Set xmlDelete = vxmlTableElement.selectSingleNode("DELETE")
    If Not xmlDelete Is Nothing Then
        For Each xmlCustomer In xmlDelete.selectNodes("CUSTOMERVERSIONBANKRUPTCYHISTORY")
            xmlCustomer.appendChild vxmlTableElement.selectSingleNode("BANKRUPTCYHISTORYGUID").cloneNode(True)
            m_objIDOAssist.Delete xmlCustomer, xmlCVBHClassDefDoc
        Next
        vxmlTableElement.removeChild xmlDelete
    End If
    'BMIDS00190 End
    
    m_objIDOAssist.Update vxmlTableElement, xmlClassDefDoc
    
    Set objICFClassDef = Nothing
    Set xmlClassDefDoc = Nothing
    'BMID000190
    Set xmlCVBHClassDefDoc = Nothing
    Set xmlCustomer = Nothing
    Set xmlDelete = Nothing
    'BMID000190 End
    
    If Not m_objContext Is Nothing Then
        m_objContext.SetComplete
    End If
    
    Exit Sub

UpdateBankruptcyHistoryVbErr:
        
    Set objICFClassDef = Nothing
    Set xmlClassDefDoc = Nothing
    'BMID000190
    Set xmlCVBHClassDefDoc = Nothing
    Set xmlCustomer = Nothing
    Set xmlDelete = Nothing
    'BMID000190 End
    
    m_objErrAssist.AddToErrSource strFunctionName
    
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

Private Sub ICustomerFinancialDO_UpdateCCJHistory(ByVal vxmlTableElement As IXMLDOMElement)
' header ----------------------------------------------------------------------------------
' description:
'   UpdateCCJHistory a single instance of the persistant data associated with this
'   data object
' pass:
'   vxmlxmlTableElement  xml element containing the request
' return:                n/a
'------------------------------------------------------------------------------------------
On Error GoTo UpdateCCJHistoryVbErr

    Const strFunctionName As String = "ICustomerFinancialDO_UpdateCCJHistory"
    
    Dim xmlClassDefDoc      As FreeThreadedDOMDocument40
    Dim objICFClassDef      As ICFClassDef
    'BMIDS00190
    Dim xmlCVCHClassDefDoc  As FreeThreadedDOMDocument40
    Dim xmlCustomer         As IXMLDOMElement
    Dim xmlDelete           As IXMLDOMNode
    'BMIDS00190 End
   
    Set objICFClassDef = New CFClassDef
    Set xmlClassDefDoc = objICFClassDef.LoadCCJHistoryData()
    
    'BMIDS00190
    Set xmlCVCHClassDefDoc = objICFClassDef.LoadCustVerCCJHistoryData()
    
    'Add new customer version links
    For Each xmlCustomer In vxmlTableElement.selectNodes("CUSTOMERVERSIONCCJHISTORY")
        xmlCustomer.appendChild vxmlTableElement.selectSingleNode("CCJHISTORYGUID").cloneNode(True)
        m_objIDOAssist.Create xmlCustomer, xmlCVCHClassDefDoc
        vxmlTableElement.removeChild xmlCustomer
    Next
    
    'Remove old customer version links
    Set xmlDelete = vxmlTableElement.selectSingleNode("DELETE")
    If Not xmlDelete Is Nothing Then
        For Each xmlCustomer In xmlDelete.selectNodes("CUSTOMERVERSIONCCJHISTORY")
            xmlCustomer.appendChild vxmlTableElement.selectSingleNode("CCJHISTORYGUID").cloneNode(True)
            m_objIDOAssist.Delete xmlCustomer, xmlCVCHClassDefDoc
        Next
        vxmlTableElement.removeChild xmlDelete
    End If
    'BMIDS00190 End
    
    m_objIDOAssist.Update vxmlTableElement, xmlClassDefDoc
    
    Set objICFClassDef = Nothing
    Set xmlClassDefDoc = Nothing
    'BMID000190
    Set xmlCVCHClassDefDoc = Nothing
    Set xmlCustomer = Nothing
    Set xmlDelete = Nothing
    'BMID000190 End
    
    If Not m_objContext Is Nothing Then
        m_objContext.SetComplete
    End If
    
    Exit Sub

UpdateCCJHistoryVbErr:
        
    Set objICFClassDef = Nothing
    Set xmlClassDefDoc = Nothing
    'BMID000190
    Set xmlCVCHClassDefDoc = Nothing
    Set xmlCustomer = Nothing
    Set xmlDelete = Nothing
    'BMID000190 End
    
    m_objErrAssist.AddToErrSource strFunctionName
    
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

Private Sub ICustomerFinancialDO_UpdateDeclinedMortgage(ByVal vxmlTableElement As IXMLDOMElement)
' header ----------------------------------------------------------------------------------
' description:
'   UpdateDeclinedMortgage a single instance of the persistant data associated with this
'   data object
' pass:
'   vxmlxmlTableElement  xml element containing the request
' return:                n/a
'------------------------------------------------------------------------------------------
On Error GoTo UpdateDeclinedMortgageVbErr

    Const strFunctionName As String = "ICustomerFinancialDO_UpdateDeclinedMortgage"
        
    Dim objErrAssist As New ErrAssist
    
    Dim objContext As ObjectContext
    Set objContext = GetObjectContext()
    
    Dim xmlClassDefDoc As FreeThreadedDOMDocument40
    Dim objICFClassDef As ICFClassDef
    Dim objIDOAssist As IDOAssist

    If objContext Is Nothing Then
      Set objIDOAssist = New DOAssist
    Else
        Set objIDOAssist = objContext.CreateInstance(gstrBASE_COMPONENT & ".DOAssist")
    End If
    
    Set objICFClassDef = New CFClassDef
    Set xmlClassDefDoc = objICFClassDef.LoadDeclinedMortgageData()
    objIDOAssist.Update vxmlTableElement, xmlClassDefDoc
    
    Set objICFClassDef = Nothing
    Set xmlClassDefDoc = Nothing
    Set objIDOAssist = Nothing
    Set objErrAssist = Nothing
    
    If Not objContext Is Nothing Then
        objContext.SetComplete
    End If
    
    Exit Sub

UpdateDeclinedMortgageVbErr:
        
    Set objICFClassDef = Nothing
    Set xmlClassDefDoc = Nothing
    Set objIDOAssist = Nothing
    
    objErrAssist.AddToErrSource strFunctionName
    
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

Private Sub ICustomerFinancialDO_UpdateGroupConnection(ByVal vxmlTableElement As IXMLDOMElement)
' header ----------------------------------------------------------------------------------
' description:
'   UpdateGroupConnection a single instance of the persistant data associated with this
'   data object
' pass:
'   vxmlxmlTableElement  xml element containing the request
' return:                n/a
'------------------------------------------------------------------------------------------
On Error GoTo UpdateGroupConnectionVbErr

    Const strFunctionName As String = "ICustomerFinancialDO_UpdateGroupConnection"
        
    Dim objErrAssist As New ErrAssist
    
    Dim objContext As ObjectContext
    Set objContext = GetObjectContext()
    
    Dim xmlClassDefDoc As FreeThreadedDOMDocument40
    Dim objICFClassDef As ICFClassDef
    Dim objIDOAssist As IDOAssist

    If objContext Is Nothing Then
       Set objIDOAssist = New DOAssist
    Else
       Set objIDOAssist = objContext.CreateInstance(gstrBASE_COMPONENT & ".DOAssist")
    End If
    
    Set objICFClassDef = New CFClassDef
    Set xmlClassDefDoc = objICFClassDef.LoadGroupConnectionData()
    objIDOAssist.Update vxmlTableElement, xmlClassDefDoc
    
    Set objICFClassDef = Nothing
    Set xmlClassDefDoc = Nothing
    Set objIDOAssist = Nothing
    Set objErrAssist = Nothing
    
    If Not objContext Is Nothing Then
        objContext.SetComplete
    End If
    
    Exit Sub

UpdateGroupConnectionVbErr:
        
    Set objICFClassDef = Nothing
    Set xmlClassDefDoc = Nothing
    Set objIDOAssist = Nothing
    
    objErrAssist.AddToErrSource strFunctionName
    
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

Private Sub ICustomerFinancialDO_UpdateLiability(ByVal vxmlTableElement As IXMLDOMElement)
' header ----------------------------------------------------------------------------------
' description:
'   UpdateLiability a single instance of the persistant data associated with this
'   data object
' pass:
'   vxmlxmlTableElement  xml element containing the request
' return:                n/a
'------------------------------------------------------------------------------------------
On Error GoTo UpdateLiabilityVbErr

    Const strFunctionName As String = "ICustomerFinancialDO_UpdateLiability"
            
    Dim xmlClassDefDoc As FreeThreadedDOMDocument40
    Dim objICFClassDef As ICFClassDef
    Dim objTPAssist As IThirdPartyAssist
    Set objTPAssist = New ThirdPartyAssist
    Dim objThirdPartyBO As ThirdPartyBO
    Dim xmlThirdPartyGUIDNode As IXMLDOMNode, _
        xmlDirectoryGUIDNode  As IXMLDOMNode
    Dim strThirdPartyGUID As String, _
        strDirectoryGUID  As String
    Dim strDeleteThirdPartyGUID As String
    Dim strRequest As String

    'BMIDS00190
    Dim xmlAccountClassDef      As FreeThreadedDOMDocument40
    Dim xmlAccountRelClassDef   As FreeThreadedDOMDocument40
    Dim xmlAccount              As IXMLDOMNode
    Dim xmlCustomer             As IXMLDOMNode
    Dim xmlDelete               As IXMLDOMNode
    Dim xmlAccountGuid          As IXMLDOMNode
    'BMIDS00190

    If m_objContext Is Nothing Then
        Set objThirdPartyBO = New ThirdPartyBO
    Else
        Set objThirdPartyBO = m_objContext.CreateInstance(gstrTHIRDPARTY_COMPONENT & ".ThirdPartyBO")
    End If
    
    Set objICFClassDef = New CFClassDef
    Set xmlClassDefDoc = objICFClassDef.LoadLoansLiabilitiesData()

    'BMIDS00190
    Set xmlAccountClassDef = objICFClassDef.LoadAccountData()
    Set xmlAccountRelClassDef = objICFClassDef.LoadAccountRelationshipData()
    Set xmlAccountGuid = vxmlTableElement.selectSingleNode("ACCOUNTGUID")
    
    'Move ThirdPartyGUID and DIRECTORYGUID one level up
    Set xmlAccount = vxmlTableElement.selectSingleNode("ACCOUNT")
    Set xmlDirectoryGUIDNode = xmlAccount.selectSingleNode("DIRECTORYGUID")
    Set xmlThirdPartyGUIDNode = xmlAccount.selectSingleNode("THIRDPARTYGUID")
    vxmlTableElement.appendChild xmlDirectoryGUIDNode
    vxmlTableElement.appendChild xmlThirdPartyGUIDNode
    
    'Add new AccountRelationships
    For Each xmlCustomer In vxmlTableElement.selectNodes("ACCOUNTRELATIONSHIP")
        xmlCustomer.appendChild xmlAccountGuid.cloneNode(True)
        m_objIDOAssist.Create xmlCustomer, xmlAccountRelClassDef
        vxmlTableElement.removeChild xmlCustomer
    Next
    
    'Remove old AccountRelationships
    Set xmlDelete = vxmlTableElement.selectSingleNode("DELETE")
    If Not xmlDelete Is Nothing Then
        For Each xmlCustomer In xmlDelete.selectNodes("ACCOUNTRELATIONSHIP")
            xmlCustomer.appendChild xmlAccountGuid.cloneNode(True)
            m_objIDOAssist.Delete xmlCustomer, xmlAccountRelClassDef
        Next
        vxmlTableElement.removeChild xmlDelete
    End If
    'BMIDS00190 End

    '
    ' Maintain the associated THIRDPARTY record if appropriate
    '
    If Not (xmlDirectoryGUIDNode Is Nothing) Then
        strDirectoryGUID = m_objXmlAssist.GetTagValue(xmlDirectoryGUIDNode, "DIRECTORYGUID")
    End If
    If Not (xmlThirdPartyGUIDNode Is Nothing) Then
        strThirdPartyGUID = m_objXmlAssist.GetTagValue(xmlThirdPartyGUIDNode, "THIRDPARTYGUID")
    End If

    If (strThirdPartyGUID <> "") And _
       ((strDirectoryGUID <> "") Or m_objXmlAssist.GetTagValue(vxmlTableElement, "THIRDPARTY") = "") Then
        ' Cannot have both a DIRECTORY and THIRDPARTY GUID or THIRDPARTYGUID and no THIRDPARTY data.
        ' Therefore, remove the THIRDPARTYGUID
        Set xmlThirdPartyGUIDNode = vxmlTableElement.selectSingleNode("THIRDPARTYGUID")
        If Not (xmlThirdPartyGUIDNode Is Nothing) Then
            xmlThirdPartyGUIDNode.Text = ""
        End If

        ' Update the primary table with the null THIRDPARTYGUID
        m_objIDOAssist.Update vxmlTableElement, xmlClassDefDoc
        
        'BMIDS00190 Restore Directory and ThirdParty GUIDs back to Account node and update
        xmlAccount.appendChild xmlDirectoryGUIDNode
        xmlAccount.appendChild xmlThirdPartyGUIDNode
        xmlAccount.appendChild xmlAccountGuid.cloneNode(True)
        m_objIDOAssist.Update xmlAccount, xmlAccountClassDef
        'BMIDS00190 End
        
        ' Delete from the THIRDPARTY table
        objTPAssist.DeleteThirdParty strThirdPartyGUID
    Else
        objTPAssist.MaintainThirdParty vxmlTableElement
        m_objIDOAssist.Update vxmlTableElement, xmlClassDefDoc
        
        'BMIDS00190 Restore Directory and ThirdParty GUIDs back to Account node and update
        xmlAccount.appendChild xmlDirectoryGUIDNode
        xmlAccount.appendChild xmlThirdPartyGUIDNode
        xmlAccount.appendChild xmlAccountGuid.cloneNode(True)
        m_objIDOAssist.Update xmlAccount, xmlAccountClassDef
        'BMIDS00190 End
    End If
    
    Set objICFClassDef = Nothing
    Set xmlClassDefDoc = Nothing
    Set objThirdPartyBO = Nothing
    Set xmlThirdPartyGUIDNode = Nothing
    Set xmlDirectoryGUIDNode = Nothing
    Set objTPAssist = Nothing
    'BMIDS00190
    Set xmlAccountClassDef = Nothing
    Set xmlAccountRelClassDef = Nothing
    Set xmlAccount = Nothing
    Set xmlCustomer = Nothing
    Set xmlDelete = Nothing
    Set xmlAccountGuid = Nothing
    'BMIDS00190 End
    
    If Not m_objContext Is Nothing Then
        m_objContext.SetComplete
    End If
    
    Exit Sub

UpdateLiabilityVbErr:
        
    Set objICFClassDef = Nothing
    Set xmlClassDefDoc = Nothing
    Set objThirdPartyBO = Nothing
    Set xmlThirdPartyGUIDNode = Nothing
    Set xmlDirectoryGUIDNode = Nothing
    Set objTPAssist = Nothing
    'BMIDS00190
    Set xmlAccountClassDef = Nothing
    Set xmlAccountRelClassDef = Nothing
    Set xmlAccount = Nothing
    Set xmlCustomer = Nothing
    Set xmlDelete = Nothing
    Set xmlAccountGuid = Nothing
    'BMIDS00190 End
    
    m_objErrAssist.AddToErrSource strFunctionName
    
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

Private Sub ICustomerFinancialDO_UpdateLifeProduct(ByVal vxmlTableElement As IXMLDOMElement, _
                                                   Optional ByVal vblnUpdatePolicyRelationship As Boolean = True)
' header ----------------------------------------------------------------------------------
' description:
'   UpdateLifeProduct a single instance of the persistant data associated with this
'   data object
' pass:
'   vxmlxmlTableElement  xml element containing the request
' return:                n/a
'------------------------------------------------------------------------------------------
On Error GoTo UpdateLifeProductVbErr

    Const strFunctionName As String = "ICustomerFinancialDO_UpdateLifeProduct"
        
    Dim objErrAssist As New ErrAssist
    
    Dim objContext As ObjectContext
    Set objContext = GetObjectContext()
    
    Dim xmlClassDefDoc As FreeThreadedDOMDocument40
    Dim objICFClassDef As ICFClassDef
    Dim objIDOAssist As IDOAssist

    Dim xmlLifePolicyNode         As IXMLDOMNode, _
        xmlACNode                 As IXMLDOMNode, _
        xmlMRCNode                As IXMLDOMNode, _
        xmlAccountNode            As IXMLDOMNode
    Dim xmlThirdPartyGUIDNode As IXMLDOMNode, _
        xmlDirectoryGUIDNode  As IXMLDOMNode
    Dim strThirdPartyGUID As String, _
        strDirectoryGUID  As String
    Dim objXmlAssist As XMLAssist, _
        objTPAssist  As IThirdPartyAssist
    Set objXmlAssist = New XMLAssist
    Set objTPAssist = New ThirdPartyAssist

    If objContext Is Nothing Then
        Set objIDOAssist = New DOAssist
    Else
        Set objIDOAssist = objContext.CreateInstance(gstrBASE_COMPONENT & ".DOAssist")
    End If
    
    Set objICFClassDef = New CFClassDef
    Set xmlLifePolicyNode = vxmlTableElement.selectSingleNode(".//LIFEPOLICY")
    If Not (xmlLifePolicyNode Is Nothing) Then
        ' Life Policy
        Set xmlClassDefDoc = objICFClassDef.LoadLifePolicyData()
        objIDOAssist.Update vxmlTableElement, xmlClassDefDoc
    End If

    Set xmlMRCNode = vxmlTableElement.selectSingleNode(".//MORTGAGERELATEDCONTRACTS")
    If Not (xmlMRCNode Is Nothing) Then
        ' Mortgage Related Contracts
        Set xmlClassDefDoc = objICFClassDef.LoadMortgageRelatedContractsData()
        objIDOAssist.Update vxmlTableElement, xmlClassDefDoc
    End If

    Set xmlACNode = vxmlTableElement.selectSingleNode(".//APPLICATIONCONTRACT")
    If Not (xmlACNode Is Nothing) Then
        ' Application Contract (New Link table)
        Set xmlClassDefDoc = objICFClassDef.LoadApplicationContractData()
        objIDOAssist.Update vxmlTableElement, xmlClassDefDoc
    End If

    Set xmlAccountNode = vxmlTableElement.selectSingleNode(".//ACCOUNT")
    If Not (xmlAccountNode Is Nothing) Then
        Set xmlClassDefDoc = objICFClassDef.LoadAccountData()

        ' Maintain the associated THIRDPARTY record if appropriate
        Set xmlDirectoryGUIDNode = xmlAccountNode.selectSingleNode("DIRECTORYGUID")
        Set xmlThirdPartyGUIDNode = xmlAccountNode.selectSingleNode("THIRDPARTYGUID")
        If Not (xmlDirectoryGUIDNode Is Nothing) Then strDirectoryGUID = objXmlAssist.GetTagValue(xmlDirectoryGUIDNode, "DIRECTORYGUID")
        If Not (xmlThirdPartyGUIDNode Is Nothing) Then strThirdPartyGUID = objXmlAssist.GetTagValue(xmlThirdPartyGUIDNode, "THIRDPARTYGUID")
    
        If (strThirdPartyGUID <> "") And _
           ((strDirectoryGUID <> "") Or objXmlAssist.GetTagValue(xmlAccountNode, "THIRDPARTY") = "") Then
            ' Cannot have both a DIRECTORY and THIRDPARTY GUID or THIRDPARTYGUID and no THIRDPARTY data.
            ' Therefore, remove the THIRDPARTYGUID
            Set xmlThirdPartyGUIDNode = xmlAccountNode.selectSingleNode("THIRDPARTYGUID")
            If Not (xmlThirdPartyGUIDNode Is Nothing) Then
                xmlThirdPartyGUIDNode.Text = ""
            End If
    
            ' Update the primary table with the null THIRDPARTYGUID
            objIDOAssist.Update xmlAccountNode, xmlClassDefDoc
            ' Delete from the THIRDPARTY table
            objTPAssist.DeleteThirdParty strThirdPartyGUID
        Else
            objTPAssist.MaintainThirdParty xmlAccountNode
            objIDOAssist.Update xmlAccountNode, xmlClassDefDoc
        End If
    End If
    
    Set objICFClassDef = Nothing
    Set xmlClassDefDoc = Nothing
    Set objIDOAssist = Nothing
    Set objErrAssist = Nothing
    Set xmlLifePolicyNode = Nothing
    Set xmlMRCNode = Nothing
    Set xmlAccountNode = Nothing
    Set objXmlAssist = Nothing
    Set xmlThirdPartyGUIDNode = Nothing
    Set xmlDirectoryGUIDNode = Nothing
    Set objTPAssist = Nothing
    
    If Not objContext Is Nothing Then
        objContext.SetComplete
    End If
    
    Exit Sub

UpdateLifeProductVbErr:
        
    Set objICFClassDef = Nothing
    Set xmlClassDefDoc = Nothing
    Set objIDOAssist = Nothing
    Set xmlLifePolicyNode = Nothing
    Set xmlMRCNode = Nothing
    Set xmlAccountNode = Nothing
    Set objXmlAssist = Nothing
    Set objTPAssist = Nothing
    
    objErrAssist.AddToErrSource strFunctionName
    
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

Private Sub ICustomerFinancialDO_UpdateMortgageAccount(ByVal vxmlTableElement As IXMLDOMElement)
' header ----------------------------------------------------------------------------------
' description:
'   UpdateMortgageAccount a single instance of the persistant data associated with this
'   data object
' pass:
'   vxmlxmlTableElement  xml element containing the request
' return:                n/a
'------------------------------------------------------------------------------------------
On Error GoTo UpdateMortgageAccountVbErr

    Const strFunctionName As String = "ICustomerFinancialDO_UpdateMortgageAccount"
        
    Dim objErrAssist As New ErrAssist
    
    Dim objContext As ObjectContext
    Set objContext = GetObjectContext()
    
    ' PSC 26/07/2002 BMIDS00006 - Start
    Dim xmlAccountRelList As IXMLDOMNode

    Dim xmlClassDefDoc As FreeThreadedDOMDocument40
    Dim objICFClassDef As ICFClassDef
    Dim objIDOAssist As IDOAssist
    Dim objXmlAssist As XMLAssist
    Set objXmlAssist = New XMLAssist
    Dim objTPAssist As IThirdPartyAssist
    Set objTPAssist = New ThirdPartyAssist
    Dim objThirdPartyBO As ThirdPartyBO
    Dim xmlThirdPartyGUIDNode As IXMLDOMNode, _
        xmlDirectoryGUIDNode  As IXMLDOMNode, _
        xmlIndemnityInsuranceNode As IXMLDOMNode
    Dim xmlAccountNode As IXMLDOMNode
    Dim strThirdPartyGUID As String, _
        strDirectoryGUID  As String
    Dim strDeleteThirdPartyGUID As String
    Dim strRequest As String

    If objContext Is Nothing Then
        Set objIDOAssist = New DOAssist
        Set objThirdPartyBO = New ThirdPartyBO
    Else
        Set objIDOAssist = objContext.CreateInstance(gstrBASE_COMPONENT & ".DOAssist")
        Set objThirdPartyBO = objContext.CreateInstance(gstrTHIRDPARTY_COMPONENT & ".ThirdPartyBO")
    End If
    
    Set objICFClassDef = New CFClassDef

    ' Mortgage Account
    Set xmlClassDefDoc = objICFClassDef.LoadMortgageAccountData()
    objIDOAssist.Update vxmlTableElement, xmlClassDefDoc

    ' Indemnity Insurance
    Set xmlIndemnityInsuranceNode = vxmlTableElement.selectSingleNode(".//INDEMNITYINSURANCE")
    If Not xmlIndemnityInsuranceNode Is Nothing Then
        Set xmlClassDefDoc = objICFClassDef.LoadIndemnityInsuranceData()
        '//MAR1488 M Heys 24/04/2006 start
        If m_objXmlAssist.GetAttributeFromNode(xmlIndemnityInsuranceNode, "CREATE") = "1" Then
            objIDOAssist.Create vxmlTableElement, xmlClassDefDoc
        Else
            objIDOAssist.Update vxmlTableElement, xmlClassDefDoc
       End If
         '//MAR1488 M Heys 24/04/2006 end
    End If

    ' Account
    Set xmlClassDefDoc = objICFClassDef.LoadAccountData()

    '
    ' Maintain the associated THIRDPARTY record if appropriate
    '
    Set xmlAccountNode = vxmlTableElement.selectSingleNode("ACCOUNT")

    If Not (xmlAccountNode Is Nothing) Then
        Set xmlDirectoryGUIDNode = xmlAccountNode.selectSingleNode("DIRECTORYGUID")
        Set xmlThirdPartyGUIDNode = xmlAccountNode.selectSingleNode("THIRDPARTYGUID")
        If Not (xmlDirectoryGUIDNode Is Nothing) Then strDirectoryGUID = objXmlAssist.GetTagValue(xmlDirectoryGUIDNode, "DIRECTORYGUID")
        If Not (xmlThirdPartyGUIDNode Is Nothing) Then strThirdPartyGUID = objXmlAssist.GetTagValue(xmlThirdPartyGUIDNode, "THIRDPARTYGUID")
    
        If (strThirdPartyGUID <> "") And _
           ((strDirectoryGUID <> "") Or objXmlAssist.GetTagValue(xmlAccountNode, "THIRDPARTY") = "") Then
            ' Cannot have both a DIRECTORY and THIRDPARTY GUID or THIRDPARTYGUID and no THIRDPARTY data.
            ' Therefore, remove the THIRDPARTYGUID
            Set xmlThirdPartyGUIDNode = xmlAccountNode.selectSingleNode("THIRDPARTYGUID")
            If Not (xmlThirdPartyGUIDNode Is Nothing) Then
                xmlThirdPartyGUIDNode.Text = ""
            End If
    
            ' Update the primary table with the null THIRDPARTYGUID
            objIDOAssist.Update xmlAccountNode, xmlClassDefDoc
            ' Delete from the THIRDPARTY table
            objTPAssist.DeleteThirdParty strThirdPartyGUID
        Else
            objTPAssist.MaintainThirdParty xmlAccountNode
            objIDOAssist.Update xmlAccountNode, xmlClassDefDoc
        End If
    End If
    
    ' PSC 26/07/2002 BMIDS00006 - Start
    Set xmlAccountRelList = m_objXmlAssist.GetMandatoryNode(vxmlTableElement, "ACCOUNTRELATIONSHIPLIST")
    ICustomerFinancialDO_CreateAccountRelationships xmlAccountRelList
    ' PSC 26/07/2002 BMIDS00006 - End
    
    Set objICFClassDef = Nothing
    Set xmlClassDefDoc = Nothing
    Set objIDOAssist = Nothing
    Set objErrAssist = Nothing
    Set objThirdPartyBO = Nothing
    Set xmlThirdPartyGUIDNode = Nothing
    Set xmlDirectoryGUIDNode = Nothing
    Set objXmlAssist = Nothing
    Set objTPAssist = Nothing
    Set xmlAccountNode = Nothing
    
    ' PSC 26/07/2002 BMIDS00006 - Start
    Set xmlAccountRelList = Nothing

    If Not objContext Is Nothing Then
        objContext.SetComplete
    End If
    
    Exit Sub

UpdateMortgageAccountVbErr:
        
    Set objICFClassDef = Nothing
    Set xmlClassDefDoc = Nothing
    Set objIDOAssist = Nothing
    Set objThirdPartyBO = Nothing
    Set xmlThirdPartyGUIDNode = Nothing
    Set xmlDirectoryGUIDNode = Nothing
    Set objXmlAssist = Nothing
    Set objTPAssist = Nothing
    Set xmlAccountNode = Nothing
    
    ' PSC 26/07/2002 BMIDS00006 - Start
    Set xmlAccountRelList = Nothing

    
    objErrAssist.AddToErrSource strFunctionName
    
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


Private Sub ICustomerFinancialDO_UpdateMortgageLoan(ByVal vxmlTableElement As IXMLDOMElement)
' header ----------------------------------------------------------------------------------
' description:
'   UpdateMortgageLoan a single instance of the persistant data associated with this
'   data object
' pass:
'   vxmlxmlTableElement  xml element containing the request
' return:                n/a
'------------------------------------------------------------------------------------------
On Error GoTo UpdateMortgageLoanVbErr

    Const strFunctionName As String = "ICustomerFinancialDO_UpdateMortgageLoan"
        
    Dim objErrAssist As New ErrAssist
    
    Dim objContext As ObjectContext
    Set objContext = GetObjectContext()
    
    Dim xmlClassDefDoc As FreeThreadedDOMDocument40
    Dim objICFClassDef As ICFClassDef
    Dim objIDOAssist As IDOAssist

    If objContext Is Nothing Then
       Set objIDOAssist = New DOAssist
    Else
       Set objIDOAssist = objContext.CreateInstance(gstrBASE_COMPONENT & ".DOAssist")
    End If
    
    Set objICFClassDef = New CFClassDef
    Set xmlClassDefDoc = objICFClassDef.LoadMortgageLoanData
    objIDOAssist.Update vxmlTableElement, xmlClassDefDoc
    
    Set objICFClassDef = Nothing
    Set xmlClassDefDoc = Nothing
    Set objIDOAssist = Nothing
    Set objErrAssist = Nothing
    
    If Not objContext Is Nothing Then
        objContext.SetComplete
    End If
    
    Exit Sub

UpdateMortgageLoanVbErr:
        
    Set objICFClassDef = Nothing
    Set xmlClassDefDoc = Nothing
    Set objIDOAssist = Nothing
    
    objErrAssist.AddToErrSource strFunctionName
    
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

Private Function ICustomerFinancialDO_FindLoansToBePorted(ByVal vxmlTableElement As IXMLDOMElement) As IXMLDOMNode
' header ----------------------------------------------------------------------------------
' description:
'   Get the data for all instances of the persistant data associated with
'   this data object for the values supplied
' pass:
'   vxmlxmlTableElement  xml element containing the request
' return:                xml node containing retrieved data
'------------------------------------------------------------------------------------------
On Error GoTo FindLoansToBePortedVbErr

    Const strFunctionName As String = "ICustomerFinancialDO_FindLoansToBePorted"

    Dim objErrAssist As New ErrAssist
    
    Dim objIADOAssist As IADOAssist, _
        objIDOAssist As IDOAssist, _
        objXmlAssist As New XMLAssist, _
        objSQLAssist As New SQLAssist
    Dim objIClassDef As ICFClassDef, _
        xmlClassDefDoc As FreeThreadedDOMDocument40
    Dim xmlReturn As New FreeThreadedDOMDocument40
    xmlReturn.validateOnParse = False
    xmlReturn.setProperty "NewParser", True
    Dim Cmd As ADODB.Command, _
        rst As ADODB.Recordset
    Dim strMortgageAccountWhere As String, _
        strMortgageLoanWhere    As String, _
        strSql                  As String
    Dim strCustomerNumber        As String, _
        strCustomerVersionNumber As String
    Dim xmlCustomerList As IXMLDOMNodeList, _
        xmlCustomer     As IXMLDOMNode
    Dim blnNeedOR As Boolean

    Dim objContext As ObjectContext
    Set objContext = GetObjectContext()
         
    Set objIDOAssist = New DOAssist
    If objContext Is Nothing Then
        Set objIADOAssist = New ADOAssist
    Else
        Set objIADOAssist = objContext.CreateInstance(gstrBASE_COMPONENT & ".ADOAssist")
    End If

    '
    ' Compile the SQL
    '
    ' MORTGAGEACCOUNT part of the SQL
    Set xmlCustomerList = vxmlTableElement.getElementsByTagName("CUSTOMER")

    blnNeedOR = False
    For Each xmlCustomer In xmlCustomerList
        strCustomerNumber = Trim$(objXmlAssist.GetTagValue(xmlCustomer, "CUSTOMERNUMBER"))
        strCustomerVersionNumber = Trim$(objXmlAssist.GetTagValue(xmlCustomer, "CUSTOMERVERSIONNUMBER"))

        If (strCustomerNumber <> "") And (strCustomerVersionNumber <> "") Then
            ' Append the CUSTOMERNUMBER and CUSTOMERVERSIONNUMBER to WHERE clause
            strMortgageAccountWhere = strMortgageAccountWhere & IIf(blnNeedOR, " OR ", "") & "(" & _
                    "CUSTOMERNUMBER = " & objSQLAssist.FormatString(strCustomerNumber) & " AND " & _
                    "CUSTOMERVERSIONNUMBER = " & strCustomerVersionNumber & ")"
            blnNeedOR = True
        End If
    Next xmlCustomer

    If strMortgageAccountWhere = "" Then
        objErrAssist.ThrowError TypeName(Me), strFunctionName, oeInvalidParameter, _
        "At least one customer must be specified"
    End If

    ' MORTGAGELOAN part of the SQL
'    strMortgageLoanWhere = "EXISTS (SELECT VALUEID FROM COMBOVALIDATION WHERE " & _
'                                    "GROUPNAME = 'RedemptionStatus' AND VALUEID = REDEMPTIONSTATUS AND " & _
'                                    "VALIDATIONTYPE = 'P') AND " & _
'                           "ACCOUNTGUID IN (SELECT ACCOUNTGUID FROM MORTGAGEACCOUNT WHERE " & _
'                                    strMortgageAccountWhere & ")"
'
'    strSql = "SELECT * FROM MORTGAGELOAN WHERE " & strMortgageLoanWhere
    
    'BMIDS00409 fixed SQL statement
    'strSql = "SELECT ML.* FROM MORTGAGELOAN ML, COMBOVALIDATION CV, ACCOUNTRELATIONSHIP AR " & _
    '         "WHERE CV.GROUPNAME = 'RedemptionStatus' AND CV.VALIDATIONTYPE = 'P' AND " & _
    '         "ML.REDEMPTIONSTATUS = CV.VALUEID AND AR.ACCOUNTGUID = ML.ACCOUNTGUID AND " & _
    '         strMortgageAccountWhere
    strSql = "SELECT  ML.* " & _
             "FROM MORTGAGELOAN ML " & _
             "INNER JOIN ACCOUNTRELATIONSHIP AR ON AR.ACCOUNTGUID = ML.ACCOUNTGUID " & _
             "INNER JOIN COMBOVALIDATION CV ON CV.VALUEID = ML.REDEMPTIONSTATUS AND CV.GROUPNAME = 'RedemptionStatus' AND CV.VALIDATIONTYPE = 'P' " & _
             "WHERE " & strMortgageAccountWhere
    'BMIDS00409 End

    '
    ' Execute the SQL
    '
    Set Cmd = New ADODB.Command
    With Cmd
        .ActiveConnection = objIADOAssist.GetConnStr
        .CommandType = adCmdText
        .CommandText = strSql

        Set rst = Cmd.Execute
    End With

    If rst.EOF Then
        ' raise application error to be interpreted by calling object
        objErrAssist.ThrowError TypeName(Me), strFunctionName, oeRecordNotFound
    End If

    '
    ' Convert the recordset into XML
    '
    Set objIClassDef = New CFClassDef
    Set xmlClassDefDoc = objIClassDef.LoadMortgageLoanData

    xmlReturn.appendChild xmlReturn.createElement("LOANSTOBEPORTED")

    ' AD 9/10/01 Removed the MoveFirst as it's a ForwardOnly cursor. SYS2718.
    Do While Not rst.EOF
        xmlReturn.documentElement.appendChild objIDOAssist.GetXMLFromRecordSet(rst, xmlClassDefDoc)
        rst.MoveNext
    Loop
        rst.Close

    Set ICustomerFinancialDO_FindLoansToBePorted = xmlReturn.documentElement

    Set objErrAssist = Nothing
    Set objIDOAssist = Nothing
    Set objIADOAssist = Nothing
    Set objXmlAssist = Nothing
    Set objSQLAssist = Nothing
    Set xmlReturn = Nothing
    Set Cmd = Nothing
    Set rst = Nothing
    Set objIClassDef = Nothing
    Set xmlClassDefDoc = Nothing
    
    If Not objContext Is Nothing Then
        objContext.SetComplete
    End If
    
    Exit Function

FindLoansToBePortedVbErr:
    
    ' Close the recordset if still open
        If Not rst Is Nothing Then
                If rst.State = adStateOpen Then
                        rst.Close
                End If
        End If

    Set objIDOAssist = Nothing
    Set objIADOAssist = Nothing
    Set objXmlAssist = Nothing
    Set objSQLAssist = Nothing
    Set xmlReturn = Nothing
    Set Cmd = Nothing
    Set rst = Nothing
    Set objIClassDef = Nothing
    Set xmlClassDefDoc = Nothing
    
    objErrAssist.AddToErrSource strFunctionName
    
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


Private Function ICustomerFinancialDO_FindLoansAndLiabilitiesWithNoLender(ByVal vxmlTableElement As IXMLDOMElement) As IXMLDOMNode
' header ----------------------------------------------------------------------------------
' description:
'   Finds the loans and liabilities which have no lender third party or directory guids
' pass:
'   vxmlxmlTableElement  xml element containing the request
' return:                xml node containing retrieved data
'------------------------------------------------------------------------------------------
On Error GoTo FindLoansAndLiabilitiesWithNoLenderVbErr

    Const strFunctionName As String = "ICustomerFinancialDO_FindLoansAndLiabilitiesWithNoLender"

    Dim objErrAssist As New ErrAssist
    
    Dim objIADOAssist As IADOAssist
    Dim objIDOAssist As IDOAssist
    Dim objXmlAssist As New XMLAssist
    Dim objSQLAssist As New SQLAssist
    Dim objIClassDef As ICFClassDef
    Dim xmlClassDefDoc As FreeThreadedDOMDocument40
    Dim xmlReturn As FreeThreadedDOMDocument40
    Dim Cmd As ADODB.Command
    Dim rst As ADODB.Recordset
    Dim strSql As String
    Dim strCustomerNumber As String
    Dim strCustomerVersionNumber As String
    Dim strLoansAndLiabilitiesWhere As String
    Dim xmlCustomerList As IXMLDOMNodeList
    Dim xmlCustomer As IXMLDOMNode
    Dim blnNeedOR As Boolean
    
    Dim objContext As ObjectContext
    Set objContext = GetObjectContext()
         
    Set objIDOAssist = New DOAssist
    If objContext Is Nothing Then
        Set objIADOAssist = New ADOAssist
    Else
        Set objIADOAssist = objContext.CreateInstance(gstrBASE_COMPONENT & ".ADOAssist")
    End If

    Set xmlCustomerList = vxmlTableElement.getElementsByTagName("LOANSLIABILITIES")
    
    blnNeedOR = False
    For Each xmlCustomer In xmlCustomerList
        strCustomerNumber = Trim$(objXmlAssist.GetTagValue(xmlCustomer, "CUSTOMERNUMBER"))
        strCustomerVersionNumber = Trim$(objXmlAssist.GetTagValue(xmlCustomer, "CUSTOMERVERSIONNUMBER"))

        If (strCustomerNumber <> "") And (strCustomerVersionNumber <> "") Then
            ' Append the CUSTOMERNUMBER and CUSTOMERVERSIONNUMBER to WHERE clause
            strLoansAndLiabilitiesWhere = strLoansAndLiabilitiesWhere & IIf(blnNeedOR, " OR ", "") & "(" & _
                    "CUSTOMERNUMBER = " & objSQLAssist.FormatString(strCustomerNumber) & " AND " & _
                    "CUSTOMERVERSIONNUMBER = " & strCustomerVersionNumber & ")"
            blnNeedOR = True
        End If
    Next xmlCustomer

    If strLoansAndLiabilitiesWhere = "" Then
        objErrAssist.ThrowError TypeName(Me), strFunctionName, oeInvalidParameter, _
        "Atleast one customer must be specified"
    End If

    ' SQL
    strLoansAndLiabilitiesWhere = "(" & strLoansAndLiabilitiesWhere & ") AND THIRDPARTYGUID is NULL AND DIRECTORYGUID is NULL " & _
        "ORDER BY CUSTOMERNUMBER, CUSTOMERVERSIONNUMBER, SEQUENCENUMBER"

    strSql = "SELECT * FROM LOANSLIABILITIES WHERE " & strLoansAndLiabilitiesWhere


    ' Execute the SQL
    Set Cmd = New ADODB.Command
    With Cmd
        .ActiveConnection = objIADOAssist.GetConnStr
        .CommandType = adCmdText
        .CommandText = strSql
        Set rst = Cmd.Execute
    End With

    If rst.EOF Then
        ' raise application error to be interpreted by calling object
        objErrAssist.ThrowError TypeName(Me), strFunctionName, oeRecordNotFound
    End If

    '
    ' Convert the recordset into XML
    '
    Set objIClassDef = New CFClassDef
    Set xmlClassDefDoc = objIClassDef.LoadLoansLiabilitiesData

    Dim xmlNode As IXMLDOMNode
    Set xmlReturn = New FreeThreadedDOMDocument40
    xmlReturn.validateOnParse = False
    xmlReturn.setProperty "NewParser", True
    xmlReturn.appendChild xmlReturn.createElement("LOANSLIABILITIESLIST")
    
    ' AD 9/10/01 Removed the MoveFirst as it's a ForwardOnly cursor. SYS2718.

    Do While Not rst.EOF
        'If xmlNode is passed in as the document owner it works under Debub but not under DLL
        'Therefore have to pass in Firstchild of return. It should point to the same thing but
        'doesn't. SYS0255
        Set xmlNode = objIDOAssist.GetXMLFromRecordSet(rst, xmlClassDefDoc, xmlReturn.firstChild)
        xmlReturn.documentElement.appendChild xmlNode
        rst.MoveNext
    Loop
        rst.Close

    Set ICustomerFinancialDO_FindLoansAndLiabilitiesWithNoLender = xmlReturn.documentElement

    Set objErrAssist = Nothing
    Set objIDOAssist = Nothing
    Set objIADOAssist = Nothing
    Set objXmlAssist = Nothing
    Set objSQLAssist = Nothing
    Set xmlReturn = Nothing
    Set Cmd = Nothing
    Set rst = Nothing
    Set objIClassDef = Nothing
    Set xmlClassDefDoc = Nothing
    
    If Not objContext Is Nothing Then
        objContext.SetComplete
    End If

    Exit Function

FindLoansAndLiabilitiesWithNoLenderVbErr:
    
    ' Close the recordset if still open
        If Not rst Is Nothing Then
                If rst.State = adStateOpen Then
                        rst.Close
                End If
        End If

    Set objIDOAssist = Nothing
    Set objIADOAssist = Nothing
    Set objXmlAssist = Nothing
    Set objSQLAssist = Nothing
    Set xmlReturn = Nothing
    Set Cmd = Nothing
    Set rst = Nothing
    Set objIClassDef = Nothing
    Set xmlClassDefDoc = Nothing
    
    objErrAssist.AddToErrSource strFunctionName
    
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

Private Function BuildRequestBlock(ByVal vxmlRequest As IXMLDOMElement, _
                                   ByVal vstrItemName As String) As IXMLDOMNode
' header ----------------------------------------------------------------------------------
' description:
'   Get the data for all instances of the persistant data associated with
'   this data object for the values supplied
' pass:
'   vxmlxmlTableElement  xml element containing the request
' return:                xml node containing retrieved data
'------------------------------------------------------------------------------------------
On Error GoTo BuildRequestBlockVbErr

    Const strFunctionName As String = "BuildRequestBlock"

    Dim objXmlAssist As XMLAssist
    Set objXmlAssist = New XMLAssist
    Dim objErrAssist As ErrAssist
    Set objErrAssist = New ErrAssist

    Dim strCustomerNumber        As String, _
        strCustomerVersionNumber As String
    Dim xmlCustomer     As IXMLDOMNode, _
        xmlCustomerList As IXMLDOMNodeList
    Dim xmlReturn      As IXMLDOMNode, _
        xmlRequestNode As IXMLDOMNode

    Set xmlReturn = vxmlRequest.ownerDocument.createElement(vstrItemName & "LIST")

    Set xmlCustomerList = vxmlRequest.getElementsByTagName("CUSTOMER")
    For Each xmlCustomer In xmlCustomerList
        strCustomerNumber = objXmlAssist.GetTagValue(xmlCustomer, "CUSTOMERNUMBER")
        strCustomerVersionNumber = objXmlAssist.GetTagValue(xmlCustomer, "CUSTOMERVERSIONNUMBER")

        Set xmlRequestNode = xmlReturn.appendChild(vxmlRequest.ownerDocument.createElement(vstrItemName))
        xmlRequestNode.appendChild(xmlRequestNode.ownerDocument.createElement("CUSTOMERNUMBER")).Text = strCustomerNumber
        xmlRequestNode.appendChild(xmlRequestNode.ownerDocument.createElement("CUSTOMERVERSIONNUMBER")).Text = strCustomerVersionNumber
    Next xmlCustomer

    Set BuildRequestBlock = xmlReturn

    Set objErrAssist = Nothing
    Set objXmlAssist = Nothing
    Set xmlCustomer = Nothing
    Set xmlCustomerList = Nothing
    Set xmlReturn = Nothing
    Set xmlRequestNode = Nothing
    
    Exit Function

BuildRequestBlockVbErr:
    Set objXmlAssist = Nothing
    Set xmlCustomer = Nothing
    Set xmlCustomerList = Nothing
    Set xmlReturn = Nothing
    Set xmlRequestNode = Nothing

    objErrAssist.AddToErrSource strFunctionName
    
    If objErrAssist.IsSystemError = True Then
        objErrAssist.LogError TypeName(Me), strFunctionName, Err.Number, Err.Description
    End If

    Set objErrAssist = Nothing
    
    '   re-raise error for business object to interpret as appropriate
    Err.Raise Err.Number, Err.Source, Err.Description

End Function

Private Sub OrderListByApplicant(ByVal vxmlRequest As IXMLDOMNode, _
                                 ByRef rxmlResults As IXMLDOMNode, _
                                 ByVal vstrItemName As String)
' header ----------------------------------------------------------------------------------
' description:
'   Get the data for all instances of the persistant data associated with
'   this data object for the values supplied
' pass:
'   vxmlxmlTableElement  xml element containing the request
' return:                xml node containing retrieved data
'------------------------------------------------------------------------------------------
On Error GoTo OrderListByApplicantVbErr

    Const strFunctionName As String = "OrderListByApplicant"

    Dim objXmlAssist As New XMLAssist
    Dim objErrAssist As New ErrAssist

    Dim xmlApplicantList As IXMLDOMNodeList, _
        xmlApplicant     As IXMLDOMNode
    Dim xmlOldNodeList As IXMLDOMNodeList, _
        xmlOldNode     As IXMLDOMNode
    Dim strCustomerNumber        As String, _
        strCustomerVersionNumber As String
    Dim xmlNewListNode As IXMLDOMNode

    Dim strListName As String
    
    strListName = rxmlResults.baseName

    ' Add a duplicate ...LIST child node to the return document to store the new list items in
    Set xmlNewListNode = rxmlResults.appendChild(rxmlResults.ownerDocument.createElement(strListName))

    'GD  SYS2514 - Set xmlOldNodeList = rxmlResults.selectNodes(vstrItemName)
    Set xmlApplicantList = vxmlRequest.selectNodes(vstrItemName)

    For Each xmlApplicant In xmlApplicantList
        ' Record the customer ID for this application
        strCustomerNumber = objXmlAssist.GetTagValue(xmlApplicant, "CUSTOMERNUMBER")
        strCustomerVersionNumber = objXmlAssist.GetTagValue(xmlApplicant, "CUSTOMERVERSIONNUMBER")

        ' Find the item in the old list corresponding to this customer ID (if any exist)
        'GD SYS2514 - Reset List For Each Applicant - XML Parser Issue
        Set xmlOldNodeList = rxmlResults.selectNodes(vstrItemName)

        For Each xmlOldNode In xmlOldNodeList
            If objXmlAssist.GetTagValue(xmlOldNode, "CUSTOMERNUMBER") = strCustomerNumber And _
               objXmlAssist.GetTagValue(xmlOldNode, "CUSTOMERVERSIONNUMBER") = strCustomerVersionNumber Then
                ' Found the item - add it to the new list
                xmlNewListNode.appendChild xmlOldNode.cloneNode(True)
            End If
        Next xmlOldNode
    Next xmlApplicant

    ' Remove all children of the current list root
    For Each xmlOldNode In xmlOldNodeList
        xmlOldNode.parentNode.removeChild xmlOldNode
    Next xmlOldNode
       
    Set rxmlResults = xmlNewListNode

    Set objErrAssist = Nothing
    Set xmlApplicantList = Nothing
    Set xmlApplicant = Nothing
    Set xmlOldNodeList = Nothing
    Set xmlOldNode = Nothing
    Set xmlNewListNode = Nothing
    
    Exit Sub

OrderListByApplicantVbErr:
    Set xmlApplicantList = Nothing
    Set xmlApplicant = Nothing
    Set xmlOldNodeList = Nothing
    Set xmlOldNode = Nothing
    Set xmlNewListNode = Nothing
    
    objErrAssist.AddToErrSource strFunctionName
    
    If objErrAssist.IsSystemError = True Then
        objErrAssist.LogError TypeName(Me), strFunctionName, Err.Number, Err.Description
    End If

    Set objErrAssist = Nothing
    
    '   re-raise error for business object to interpret as appropriate
    Err.Raise Err.Number, Err.Source, Err.Description

End Sub

Private Sub GetMortgageLoanBalance(objXMLReturnNode As IXMLDOMNode)
' header ----------------------------------------------------------------------------------
' description:
'   Retrieve total of loans not being redeemed and append a new child element to the
'   relevant MortgageAccount node in the main XML block being returned
' pass:
'   objXMLReturnNode     xml node containing the existing return XML
'
'15/05/2003 LDM BM0492   If extra paramenter "SHOWNOTREDEEMED" has been tacked onto the end of the XML input.
'                        If set to true, return outstanding balance of total loans not redeemed
'-----------------------------------------------------------------------------------------
On Error GoTo GetMortgageLoanBalance_VbErr

Const strFunctionName As String = "GetMortgageLoanBalance"

Dim conn As ADODB.Connection
Dim rst As ADODB.Recordset
Dim strSql As String
Dim strAccountGUID As String

Dim objXmlNodeList As IXMLDOMNodeList
Dim objXmlNode As IXMLDOMNode
Dim objXMLElement As IXMLDOMElement
Dim objXmlChildNode As IXMLDOMNode

Dim objXmlAssist As XMLAssist
Dim objSQLAssist As SQLAssist
Dim objADOAssist As ADOAssist
Dim objErrAssist As ErrAssist

Dim dblTotalOutstanding As Double
Dim xmlShowNotRedeemedNode As IXMLDOMNode 'BM0492 LDM 15/05/2003
Dim bShowNotRedeemedNode As Boolean       ' //

    bShowNotRedeemedNode = False          'BM0492 LDM 15/05/2003
    Set objXmlAssist = New XMLAssist
    Set objSQLAssist = New SQLAssist
    Set objADOAssist = New ADOAssist
    Set objErrAssist = New ErrAssist
    
    strSql = "SELECT AccountGUID, Sum(OutstandingBalance) As TotalOutstanding FROM MortgageLoan WHERE AccountGUID IN ("
    
    'Retrieve each MortgageAccount node...
    Set objXmlNodeList = objXMLReturnNode.selectNodes("MORTGAGEACCOUNT")
    If Not (objXmlNodeList Is Nothing) Then
        If objXmlNodeList.length > 0 Then
            For Each objXmlNode In objXmlNodeList
                '... and aa account guids to SQL string
                strAccountGUID = objXmlAssist.GetElementText(objXmlNode, "ACCOUNTGUID")
                'ASt 09/03/01 CC012 Now compatible with SQL Server and Oracle
#If GENERIC_SQL Then
                strAccountGUID = objSQLAssist.FormatGuid(strAccountGUID)
#Else
                strAccountGUID = objSQLAssist.FormatString(strAccountGUID)
#End If
                strSql = strSql & strAccountGUID & ", "
            Next
            strSql = Left$(strSql, Len(strSql) - 2) & ") " 'Remove trailing comma and space
            
            'BM0492 LDM 15/05/2003 start
            Set xmlShowNotRedeemedNode = m_objXmlAssist.GetNode(objXMLReturnNode, "SHOWNOTREDEEMED", False)
            If Not (xmlShowNotRedeemedNode Is Nothing) Then
                bShowNotRedeemedNode = CBool(xmlShowNotRedeemedNode.Text)
            End If

            If bShowNotRedeemedNode Then
                'Only retrieve amounts where loans are not redeemed
                strSql = strSql & "AND RedemptionStatus IN (SELECT VALUEID FROM COMBOVALIDATION WHERE GROUPNAME = 'RedemptionStatus' AND VALIDATIONTYPE NOT IN ('A'))"
            Else
                'Only retrieve amounts where not being/been redeemed
                strSql = strSql & "AND RedemptionStatus IN (SELECT VALUEID FROM COMBOVALIDATION WHERE GROUPNAME = 'RedemptionStatus' AND VALIDATIONTYPE IN ('N', 'P'))"
            End If
            'BM0492 LDM 15/05/2003 end
            
            strSql = strSql & " GROUP BY AccountGUID"
            
            'Get Database Connection
            Set conn = New ADODB.Connection
            conn.ConnectionString = objADOAssist.GetConnStr
            conn.open
                
            'Open recordset
            Set rst = New ADODB.Recordset
            rst.CursorLocation = adUseClient
            rst.open strSql, conn, adOpenStatic, adLockReadOnly, adCmdText
            
            'Disconnect recordset
            Set rst.ActiveConnection = Nothing

    
            If Not rst.EOF Then
                'Add new child item to each MortgageAccount node to hold the total outstanding balance
                For Each objXmlNode In objXmlNodeList
                    strAccountGUID = objXmlAssist.GetElementText(objXmlNode, "ACCOUNTGUID")
                    strAccountGUID = objSQLAssist.FormatString(strAccountGUID)
                    rst.MoveFirst   'Need to MoveFirst to search from beginning of recordset!
                    rst.Find "AccountGuid = " & strAccountGUID
                    If Not rst.EOF Then
                        'Ensure TotalOutstanding is not null!
                        If IsNull(rst("TotalOutstanding")) Then
                            dblTotalOutstanding = 0
                        Else
                            dblTotalOutstanding = rst("TotalOutstanding")
                        End If
                    Else
                        dblTotalOutstanding = 0
                    End If
                    Set objXMLElement = objXmlNode.ownerDocument.createElement("OUTSTANDINGBALANCE")
                    Set objXmlChildNode = objXmlNode.appendChild(objXMLElement)
                    objXmlChildNode.Text = dblTotalOutstanding
                Next
            End If
        End If
    End If
        rst.Close
        conn.Close
GetMortgageLoanBalance_Exit:
    Set conn = Nothing
    Set rst = Nothing
    Set objXmlNodeList = Nothing
    Set objXmlNode = Nothing
    Set objXMLElement = Nothing
    Set objXmlChildNode = Nothing
    Set objXmlAssist = Nothing
    Set objSQLAssist = Nothing
    Set objADOAssist = Nothing
    Set objErrAssist = Nothing
    Set xmlShowNotRedeemedNode = Nothing 'BM0492 LDM 15/05/2003
    Exit Sub

GetMortgageLoanBalance_VbErr:

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

    Set conn = Nothing
    Set rst = Nothing
    Set objXmlNodeList = Nothing
    Set objXmlNode = Nothing
    Set objXMLElement = Nothing
    Set objXmlChildNode = Nothing
    Set objXmlAssist = Nothing
    Set objSQLAssist = Nothing
    Set objADOAssist = Nothing
    Set xmlShowNotRedeemedNode = Nothing 'BM0492 LDM 15/05/2003
    
    'MS 27/06/99 Re-raise error correctly
    
    objErrAssist.AddToErrSource strFunctionName
    
    If objErrAssist.IsSystemError = True Then
        objErrAssist.LogError TypeName(Me), strFunctionName, Err.Number, Err.Description
    End If
    
'    If Not objContext Is Nothing Then
'        If objErrAssist.IsSystemError Then
'            objContext.SetAbort
'        Else
'            objContext.SetComplete
'        End If
'    End If

    Set objErrAssist = Nothing
    '   re-raise error for business object to interpret as appropriate
    Err.Raise Err.Number, Err.Source, Err.Description
    
'    objErrAssist.ThrowError TypeName(Me), strFunctionName, oeUnspecifiedError, "Error retrieving 'Total Loans Not Being Redeemed'"
    
End Sub

'BMIDS00403
Private Sub ICustomerFinancialDO_UpdateRegularOutgoings(ByVal vxmlTableElement As IXMLDOMElement)
' header ----------------------------------------------------------------------------------
' description:
'   UpdateRegularOutgoings a single instance of the persistant data associated with this
'   data object
' pass:
'   vxmlxmlTableElement  xml element containing the request
' return:                n/a
'------------------------------------------------------------------------------------------
On Error GoTo UpdateRegularOutgoingsVbErr

    Const strFunctionName As String = "ICustomerFinancialDO_UpdateRegularOutgoings"
            
    Dim xmlClassDefDoc      As FreeThreadedDOMDocument40
    Dim objICFClassDef      As ICFClassDef
    'BMIDS00190
    Dim xmlCVROClassDefDoc  As FreeThreadedDOMDocument40
    Dim xmlCustomer         As IXMLDOMElement
    Dim xmlDelete           As IXMLDOMNode
    'BMIDS00190 End
    
    Set objICFClassDef = New CFClassDef
    Set xmlClassDefDoc = objICFClassDef.LoadRegularOutgoingsData()
    
    'BMIDS00190
    Set xmlCVROClassDefDoc = objICFClassDef.LoadCustVerRegularOutgoingsData()
    
    'Add new customer version links
    For Each xmlCustomer In vxmlTableElement.selectNodes("CUSTOMERVERSIONREGULAROUTGOINGS")
        xmlCustomer.appendChild vxmlTableElement.selectSingleNode("REGULAROUTGOINGSGUID").cloneNode(True)
        m_objIDOAssist.Create xmlCustomer, xmlCVROClassDefDoc
        vxmlTableElement.removeChild xmlCustomer
    Next
    
    'Remove old customer version links
    Set xmlDelete = vxmlTableElement.selectSingleNode("DELETE")
    If Not xmlDelete Is Nothing Then
        For Each xmlCustomer In xmlDelete.selectNodes("CUSTOMERVERSIONREGULAROUTGOINGS")
            xmlCustomer.appendChild vxmlTableElement.selectSingleNode("REGULAROUTGOINGSGUID").cloneNode(True)
            m_objIDOAssist.Delete xmlCustomer, xmlCVROClassDefDoc
        Next
        vxmlTableElement.removeChild xmlDelete
    End If
    'BMIDS00190 End
    
    m_objIDOAssist.Update vxmlTableElement, xmlClassDefDoc

    Set objICFClassDef = Nothing
    Set xmlClassDefDoc = Nothing
    'BMID000190
    Set xmlCVROClassDefDoc = Nothing
    Set xmlCustomer = Nothing
    Set xmlDelete = Nothing
    'BMID000190 End

    If Not m_objContext Is Nothing Then
        m_objContext.SetComplete
    End If

    Exit Sub

UpdateRegularOutgoingsVbErr:
        
    Set objICFClassDef = Nothing
    Set xmlClassDefDoc = Nothing
    'BMID000190
    Set xmlCVROClassDefDoc = Nothing
    Set xmlCustomer = Nothing
    Set xmlDelete = Nothing
    'BMID000190 End
    
    m_objErrAssist.AddToErrSource strFunctionName
    
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
'BMIDS00403 End

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

Private Sub AppendContactTelephoneData(ByVal vxmlContactDetailsNode As IXMLDOMNode)
On Error GoTo AppendContactTelephoneDataVbErr

    Const strFunctionName As String = "AppendContactTelephoneData"

    Dim objErrAssist As New ErrAssist
    
    Dim objContext As ObjectContext
    Set objContext = GetObjectContext()
    
    Dim objIContactBO As IContactBO
    Dim xmlCTNode As IXMLDOMNode
    Dim xmlCTResponseNode As IXMLDOMNode
    Dim xmlContactNode As IXMLDOMNode
    Dim xmlCTNodeList As IXMLDOMNodeList
        
    Set objIContactBO = New ContactBO
    
    'Check for ContactDetailsGuid
    If m_objXmlAssist.GetElementText(vxmlContactDetailsNode, "CONTACTDETAILSGUID") <> "" Then
        Set xmlCTResponseNode = objIContactBO.GetContact(vxmlContactDetailsNode)
   
        'Append Contact Telephone data
        If Not xmlCTResponseNode Is Nothing Then
            Set xmlCTNodeList = xmlCTResponseNode.selectNodes(".//CONTACTTELEPHONEDETAILS")
    
            'Now append to ContactDetails Node
            For Each xmlCTNode In xmlCTNodeList
                vxmlContactDetailsNode.appendChild xmlCTNode
            Next
        End If
    End If

    Set objErrAssist = Nothing
    Set objIContactBO = Nothing
    Set xmlCTNode = Nothing
    Set xmlCTResponseNode = Nothing
    Set xmlContactNode = Nothing
    Set xmlCTNodeList = Nothing

    Exit Sub

AppendContactTelephoneDataVbErr:

    Set objIContactBO = Nothing
    Set xmlCTNode = Nothing
    Set xmlCTResponseNode = Nothing
    Set xmlContactNode = Nothing
    Set xmlCTNodeList = Nothing

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

Private Sub ICustomerFinancialDO_CreateAccountRelationships(ByVal vxmlTableElement As IXMLDOMElement)
' header ----------------------------------------------------------------------------------
' description: Create instances of Account Relationship
' pass:
'   vxmlTableElement  xml element containing the request
' return:                n/a
'------------------------------------------------------------------------------------------
On Error GoTo CreateAccountRelationshipsVbErr

    Const strFunctionName As String = "ICustomerFinancialDO_CreateAccountRelationships"
    

    Dim xmlClassDefDoc As FreeThreadedDOMDocument40
    Dim objIClassDef As ICFClassDef
    Dim xmlAccountRelList As IXMLDOMNodeList
    Dim xmlNode As IXMLDOMNode
    Dim xmlAccountRel As IXMLDOMNode
    
    Dim lngErrNo As Long
    Dim strDescription As String
    Dim strSource As String
                                       
    Set objIClassDef = New CFClassDef
    Set xmlClassDefDoc = objIClassDef.LoadAccountRelationshipData()
    
    Set xmlAccountRelList = vxmlTableElement.selectNodes("ACCOUNTRELATIONSHIP")
    
    If xmlAccountRelList.length > 0 Then
        
        ' Delete existing relationships
        Set xmlAccountRel = xmlClassDefDoc.createElement("ACCOUNTRELATIONSHIP")
        Set xmlNode = m_objXmlAssist.GetMandatoryNode(xmlAccountRelList.Item(0), "ACCOUNTGUID")
        xmlAccountRel.appendChild xmlNode.cloneNode(True)

        On Error Resume Next
        ICustomerFinancialDO_DeleteAccountRelationships xmlAccountRel
        
        lngErrNo = Err.Number
        strDescription = Err.Description
        strSource = Err.Source
        
        On Error GoTo CreateAccountRelationshipsVbErr
        
        ' May not be any rows to delete
        If lngErrNo <> 0 And m_objErrAssist.GetOmigaErrorNumber(lngErrNo) <> oeNoRowsAffectedByDeleteAll Then
            Err.Raise lngErrNo, strSource, strDescription
        End If
        
        ' Create new relationships
        For Each xmlNode In xmlAccountRelList
            m_objIDOAssist.Create xmlNode, xmlClassDefDoc
        Next
    Else
        m_objErrAssist.ThrowError TypeName(Me), strFunctionName, oeMissingPrimaryTag, "ACCOUNTRELATIONSHIP"
    End If
    
    Set objIClassDef = Nothing
    Set xmlClassDefDoc = Nothing
    Set xmlAccountRelList = Nothing
    Set xmlNode = Nothing
    
    m_objContext.SetComplete


    Exit Sub

CreateAccountRelationshipsVbErr:
    

    Set objIClassDef = Nothing
    Set xmlClassDefDoc = Nothing
    Set xmlAccountRelList = Nothing
    Set xmlNode = Nothing
    
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

Private Sub ICustomerFinancialDO_DeleteAccountRelationships(ByVal vxmlTableElement As IXMLDOMElement)
' header ----------------------------------------------------------------------------------
' description: Delete instances of Account Relationship
' pass:
'   vxmlTableElement  xml element containing the request
' return:                n/a
'------------------------------------------------------------------------------------------
On Error GoTo DeleteAccountRelationshipsVbErr

    Const strFunctionName As String = "ICustomerFinancialDO_DeleteAccountRelationships"
    

    Dim xmlClassDefDoc As FreeThreadedDOMDocument40
    Dim objIClassDef As ICFClassDef
    
    ' Check we have an account guid
    Dim strAccountGUID As String
    
    strAccountGUID = m_objXmlAssist.GetMandatoryElementText(vxmlTableElement, "ACCOUNTGUID")
                                       
    Set objIClassDef = New CFClassDef
    Set xmlClassDefDoc = objIClassDef.LoadAccountRelationshipData()
    m_objIDOAssist.DeleteAll vxmlTableElement, xmlClassDefDoc
    
    Set objIClassDef = Nothing
    Set xmlClassDefDoc = Nothing
    
    m_objContext.SetComplete


    Exit Sub

DeleteAccountRelationshipsVbErr:
    

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

Private Sub ICustomerFinancialDO_UpdateAccountRelationships(ByVal vxmlTableElement As IXMLDOMElement)
' header ----------------------------------------------------------------------------------
' description: Update instances of Account Relationship
' pass:
'   vxmlTableElement  xml element containing the request
' return:                n/a
'------------------------------------------------------------------------------------------
On Error GoTo UpdateAccountRelationshipsVbErr

    Const strFunctionName As String = "ICustomerFinancialDO_UpdateAccountRelationships"
    

    Dim xmlClassDefDoc As FreeThreadedDOMDocument40
    Dim objIClassDef As ICFClassDef
                                       
    Set objIClassDef = New CFClassDef
    Set xmlClassDefDoc = objIClassDef.LoadAccountRelationshipData()
    m_objIDOAssist.Update vxmlTableElement, xmlClassDefDoc
    
    Set objIClassDef = Nothing
    Set xmlClassDefDoc = Nothing

    m_objContext.SetComplete


    Exit Sub

UpdateAccountRelationshipsVbErr:
    

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

Private Function ICustomerFinancialDO_FindAccountRelationshipList(ByVal vxmlTableElement As IXMLDOMElement) As IXMLDOMNode
    
    On Error GoTo FindAccountRelationshipListVbErr

    Const strFunctionName As String = "ICustomerFinancialDO_FindAccountRelationshipList"
    
    
    Dim objIADOAssist  As IADOAssist
    Dim objSQLAssist As SQLAssist
    Dim objXmlAssist As New XMLAssist
    
    Dim adoConnection As ADODB.Connection
    Dim adoCommand As ADODB.Command
    Dim adoParameter As ADODB.Parameter
    Dim rs As ADODB.Recordset
    
    Dim xmlClassDefDoc As FreeThreadedDOMDocument40
    Dim objIClassDef As ICFClassDef
    Dim xmlData As IXMLDOMNode
    Dim xmlCopy As IXMLDOMNode
    Dim xmlOut As New FreeThreadedDOMDocument40
    xmlOut.validateOnParse = False
    xmlOut.setProperty "NewParser", True
    Dim xmlTempElement  As IXMLDOMElement
    
    Dim strAccountGUID  As String
    Dim strCustomerNumber  As String
    Dim strCustomerVersionNumber As String
    Dim strSql  As String
    
    Set objSQLAssist = New SQLAssist
    Set adoConnection = New ADODB.Connection
    Set adoCommand = New ADODB.Command
    Set objIClassDef = New CFClassDef
    Set rs = New ADODB.Recordset
    
    Set objIADOAssist = m_objContext.CreateInstance(gstrBASE_COMPONENT & ".ADOAssist")
    Set xmlClassDefDoc = objIClassDef.LoadAccountRelationshipDetailsData()
    
    Set xmlCopy = vxmlTableElement.cloneNode(True)
    
    strAccountGUID = objXmlAssist.GetTagValue(xmlCopy, "ACCOUNTGUID", True)
    strCustomerNumber = Trim$(objXmlAssist.GetTagValue(xmlCopy, "CUSTOMERNUMBER", False))
    strCustomerVersionNumber = Trim$(objXmlAssist.GetTagValue(xmlCopy, "CUSTOMERVERSIONNUMBER", False))
    
    If Len(strAccountGUID) = 0 Then
        m_objErrAssist.ThrowError TypeName(Me), strFunctionName, oeInvalidParameter, _
                        "Atleast AccountGUID must be specified"
    End If
    
    
    #If GENERIC_SQL Then
        Dim eDbEngineType As DBENGINETYPE
        eDbEngineType = objIADOAssist.GetDBEngine
        If eDbEngineType = SQLServer Then
            strSql = "USP_ACCOUNTRELATIONSHIPDETAILS"
        End If
    #End If
    
    With adoConnection
        .ConnectionString = objIADOAssist.GetConnStr
        #If Not GENERIC_SQL Then
            .CursorLocation = adUseClient
        #End If
        .open
    End With
    
    Set adoCommand.ActiveConnection = adoConnection
    adoCommand.CommandType = adCmdStoredProc
    adoCommand.CommandText = strSql
    
    'AccountGUID
    Set adoParameter = adoCommand.CreateParameter("ACCOUNTGUID", adBinary, adParamInput, 16, objSQLAssist.GuidStringToByteArray(strAccountGUID))
    adoCommand.Parameters.Append adoParameter
    If Len(strAccountGUID) > 0 Then
        adoCommand(0) = objSQLAssist.GuidStringToByteArray(strAccountGUID)
    Else
        adoCommand(0) = Null
    End If
    
    'Customer Number
    Set adoParameter = adoCommand.CreateParameter("CUSTOMERNUMBER", adBSTR, adParamInput, Len(strCustomerNumber))
    adoParameter.Attributes = adParamNullable
    adoCommand.Parameters.Append adoParameter
    If Len(strCustomerNumber) > 0 Then
        adoCommand(1) = strCustomerNumber
    Else
        adoCommand(1) = Null
    End If
    
    ' Customer Version Number
    Set adoParameter = adoCommand.CreateParameter("CUSTOMERVERSIONNUMBER", adInteger, adParamInput, Len(strCustomerVersionNumber))
    adoParameter.Attributes = adParamNullable
    adoCommand.Parameters.Append adoParameter
    If Len(strCustomerVersionNumber) > 0 Then
        adoCommand(2) = strCustomerVersionNumber
    Else
        adoCommand(2) = Null
    End If
    
    Set rs = adoCommand.Execute
    
    Set xmlTempElement = xmlOut.createElement("ACCOUNTRELATIONSHIPLIST")
    
    While Not rs.EOF
        Set xmlData = m_objIDOAssist.GetXMLFromRecordSet(rs, xmlClassDefDoc)
        m_objXmlAssist.ChangeNodeName xmlData, "ACCOUNTRELATIONSHIPDETAILS", "ACCOUNTRELATIONSHIP"
        xmlTempElement.appendChild xmlData
        rs.MoveNext
    Wend
    
    Set ICustomerFinancialDO_FindAccountRelationshipList = xmlTempElement

    Set objIClassDef = Nothing
    Set xmlClassDefDoc = Nothing
    Set xmlData = Nothing
    Set xmlCopy = Nothing
    Set objIADOAssist = Nothing
    Set objSQLAssist = Nothing
    Set objXmlAssist = Nothing
    Set adoConnection = Nothing
    Set adoCommand = Nothing
    Set adoParameter = Nothing
    Set rs = Nothing
    Set xmlOut = Nothing
    Set xmlTempElement = Nothing
    
    m_objContext.SetComplete
    
    
    Exit Function

FindAccountRelationshipListVbErr:
    
    
    Set objIClassDef = Nothing
    Set xmlClassDefDoc = Nothing
    Set xmlData = Nothing
    Set xmlCopy = Nothing
    Set objIADOAssist = Nothing
    Set objSQLAssist = Nothing
    Set objXmlAssist = Nothing
    Set adoConnection = Nothing
    Set adoCommand = Nothing
    Set adoParameter = Nothing
    Set rs = Nothing
    Set xmlOut = Nothing
    Set xmlTempElement = Nothing
    
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


Private Sub ICustomerFinancialDO_CreateSpecialFeature(ByVal vxmlTableElement As IXMLDOMElement, ByVal vxmlKeys As IXMLDOMElement)
' header ----------------------------------------------------------------------------------
' description: Create instances of Mortgage Account Special Feature
' pass:
'   vxmlTableElement  xml element containing the request
' return:                n/a
'------------------------------------------------------------------------------------------
On Error GoTo CreateSpecialFeatureVbErr

    Const strFunctionName As String = "ICustomerFinancialDO_CreateSpecialFeature"
    

    Dim xmlClassDefDoc As FreeThreadedDOMDocument40
    Dim objIClassDef As ICFClassDef
    Dim xmlSeqNo As IXMLDOMNode
    
    Set objIClassDef = New CFClassDef
    Set xmlClassDefDoc = objIClassDef.LoadMortgageAccountSpecialFeatureData()

    Dim blnGeneratedKeys As Boolean
    
    Set xmlSeqNo = m_objXmlAssist.GetNode(vxmlTableElement, "MORTGAGEACCOUNTSPECIALFEATURESEQNO", False)
    
    blnGeneratedKeys = False
    
    If xmlSeqNo Is Nothing Then
        blnGeneratedKeys = True
    ElseIf Len(xmlSeqNo.Text) = 0 Then
        blnGeneratedKeys = True
    End If
        
    If blnGeneratedKeys Then
        ' Add in the sequence number
        m_objIDOAssist.GenerateSequenceNumber vxmlTableElement, xmlClassDefDoc, "MORTGAGEACCOUNTSPECIALFEATURESEQNO"
    
        Set xmlSeqNo = m_objXmlAssist.GetNode(vxmlTableElement, "MORTGAGEACCOUNTSPECIALFEATURESEQNO")
        vxmlKeys.appendChild xmlSeqNo.cloneNode(True)
    End If
    
    m_objIDOAssist.Create vxmlTableElement, xmlClassDefDoc
    
    Set objIClassDef = Nothing
    Set xmlClassDefDoc = Nothing

    m_objContext.SetComplete


    Exit Sub

CreateSpecialFeatureVbErr:
    

    Set objIClassDef = Nothing
    Set xmlClassDefDoc = Nothing
    Set xmlSeqNo = Nothing
    
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

Private Sub ICustomerFinancialDO_DeleteSpecialFeature(ByVal vxmlTableElement As IXMLDOMElement)
' header ----------------------------------------------------------------------------------
' description: Delete instances of Mortgage Account Special Feature
' pass:
'   vxmlTableElement  xml element containing the request
' return:                n/a
'------------------------------------------------------------------------------------------
On Error GoTo DeleteSpecialFeatureVbErr

    Const strFunctionName As String = "ICustomerFinancialDO_DeleteSpecialFeature"
    
                   
    Dim xmlClassDefDoc As FreeThreadedDOMDocument40
    Dim objIClassDef As ICFClassDef
                
    Set objIClassDef = New CFClassDef
    Set xmlClassDefDoc = objIClassDef.LoadMortgageAccountSpecialFeatureData()
    m_objIDOAssist.Delete vxmlTableElement, xmlClassDefDoc
    
    Set objIClassDef = Nothing
    Set xmlClassDefDoc = Nothing
    
    m_objContext.SetComplete
    
    
    Exit Sub

DeleteSpecialFeatureVbErr:

    
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

Private Sub ICustomerFinancialDO_UpdateSpecialFeature(ByVal vxmlTableElement As IXMLDOMElement)
' header ----------------------------------------------------------------------------------
' description: Update instances of Mortgage Account Special Feature
' pass:
'   vxmlTableElement  xml element containing the request
' return:                n/a
'------------------------------------------------------------------------------------------
On Error GoTo UpdateSpecialFeatureVbErr

    Const strFunctionName As String = "ICustomerFinancialDO_UpdateSpecialFeature"
    
    
    Dim xmlClassDefDoc As FreeThreadedDOMDocument40
    Dim objIClassDef As ICFClassDef
    
    Set objIClassDef = New CFClassDef
    Set xmlClassDefDoc = objIClassDef.LoadMortgageAccountSpecialFeatureData()
    m_objIDOAssist.Update vxmlTableElement, xmlClassDefDoc
    
    Set objIClassDef = Nothing
    Set xmlClassDefDoc = Nothing
    
    m_objContext.SetComplete
        
    
    Exit Sub

UpdateSpecialFeatureVbErr:
            
    
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

Private Function ICustomerFinancialDO_FindSpecialFeatureList(ByVal vxmlTableElement As IXMLDOMElement) As IXMLDOMNode
' header ----------------------------------------------------------------------------------
' description: Find instances of Mortgage Account Special Feature
' pass:
'   vxmlTableElement  xml element containing the request
' return:                xml node containing retrieved data
'------------------------------------------------------------------------------------------
On Error GoTo FindSpecialFeatureListVbErr

    Const strFunctionName As String = "ICustomerFinancialDO_FindSpecialFeatureList"
    
    
    Dim xmlClassDefDoc As FreeThreadedDOMDocument40
    Dim objIClassDef As ICFClassDef
         
    Set objIClassDef = New CFClassDef
    Set xmlClassDefDoc = objIClassDef.LoadMortgageAccountSpecialFeatureData()
    Set ICustomerFinancialDO_FindSpecialFeatureList = m_objIDOAssist.FindList(vxmlTableElement, xmlClassDefDoc)
    Set objIClassDef = Nothing
    Set xmlClassDefDoc = Nothing
    
    m_objContext.SetComplete
    
    
    Exit Function

FindSpecialFeatureListVbErr:
    
    
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


Private Function ICustomerFinancialDO_FindMortgageAccountList(ByVal vxmlTableElement As IXMLDOMElement) As IXMLDOMNode
' header ----------------------------------------------------------------------------------
' description: Find instances of Account Relationship
' pass:
'   vxmlTableElement  xml element containing the request
' return:                n/a
'15/05/2003 LDM BM0492 if extra paramenter "SHOWNOTREDEEMED" has been tacked onto the end of the XML input
'                               make sure it is passed down to GetMortgageLoanBalance
'------------------------------------------------------------------------------------------
On Error GoTo FindMortgageAccountListVbErr

    Const strFunctionName As String = "ICustomerFinancialDO_FindMortgageAccountList"
    
    
    Dim xmlClassDefDoc As FreeThreadedDOMDocument40
    Dim objIClassDef As ICFClassDef
    Dim xmlData As IXMLDOMNode
    Dim xmlCopyNode As IXMLDOMNode
    Dim xmlAccountRelation As IXMLDOMNode
    Dim xmlAccountGuid As IXMLDOMNode
    Dim xmlAccount As IXMLDOMNode
    Dim xmlAccountList As IXMLDOMNodeList
    Dim xmlRelData As IXMLDOMNode
    'BMIDS00405
    Dim xmlResults          As IXMLDOMNode
    Dim xmlAccountExists    As IXMLDOMNode
    'BMIDS00405 End
    Dim xmlShowNotRedeemedNode As IXMLDOMNode 'BM0492 LDM 15/05/2003
    
    Set xmlAccountRelation = vxmlTableElement.ownerDocument.createElement("ACCOUNTRELATIONSHIP")
    Set xmlAccountGuid = vxmlTableElement.ownerDocument.createElement("ACCOUNTGUID")
    xmlAccountRelation.appendChild xmlAccountGuid
    
         
    Set objIClassDef = New CFClassDef
    Set xmlClassDefDoc = objIClassDef.LoadMortgageAccountListData()
    
    Set xmlCopyNode = vxmlTableElement.cloneNode(True)
    m_objXmlAssist.ChangeNodeName xmlCopyNode, "MORTGAGEACCOUNT", "MORTGAGEACCOUNTSUMMARY"
    
    Set xmlData = m_objIDOAssist.FindListMultiple(xmlCopyNode, xmlClassDefDoc)
    
    m_objXmlAssist.ChangeNodeName xmlData, "MORTGAGEACCOUNTSUMMARYLIST", "MORTGAGEACCOUNTLIST"
    m_objXmlAssist.ChangeNodeName xmlData, "MORTGAGEACCOUNTSUMMARY", "MORTGAGEACCOUNT"
    
    'BM0492 LDM 15/05/2003
    Set xmlShowNotRedeemedNode = m_objXmlAssist.GetNode(xmlCopyNode, "SHOWNOTREDEEMED", False)
    If Not (xmlShowNotRedeemedNode Is Nothing) Then
        xmlData.appendChild xmlShowNotRedeemedNode
    End If

    
    GetMortgageLoanBalance xmlData
    
    Set xmlAccountList = xmlData.selectNodes("MORTGAGEACCOUNT")
    Set xmlResults = xmlClassDefDoc.createElement("MORTGAGEACCOUNTLIST")  'BMIDS00405
    
    For Each xmlAccount In xmlAccountList
        xmlAccountGuid.Text = m_objXmlAssist.GetMandatoryElementText(xmlAccount, "ACCOUNTGUID")
        'BMIDS00405 Add the account to results if it is not already there
        Set xmlAccountExists = xmlResults.selectSingleNode("MORTGAGEACCOUNT[ACCOUNTGUID[.='" & xmlAccountGuid.Text & "']]")
        If xmlAccountExists Is Nothing Then
            xmlResults.appendChild xmlAccount
        'BMIDS00405 End
            Set xmlRelData = ICustomerFinancialDO_FindAccountRelationshipList(xmlAccountRelation)
            xmlAccount.appendChild xmlRelData
        End If 'BMIDS00405
    Next
    
    Set ICustomerFinancialDO_FindMortgageAccountList = xmlResults
    
    Set objIClassDef = Nothing
    Set xmlClassDefDoc = Nothing
    Set xmlData = Nothing
    Set xmlCopyNode = Nothing
    Set xmlAccountRelation = Nothing
    Set xmlAccountGuid = Nothing
    Set xmlAccount = Nothing
    Set xmlAccountList = Nothing
    Set xmlRelData = Nothing
    'BMIDS00405
    Set xmlResults = Nothing
    Set xmlAccountExists = Nothing
    'BMIDS00405 End
    Set xmlShowNotRedeemedNode = Nothing 'BM0492 LDM 15/05/2003
    
    m_objContext.SetComplete
    
    
    Exit Function

FindMortgageAccountListVbErr:
    
    
    Set objIClassDef = Nothing
    Set xmlClassDefDoc = Nothing
    Set xmlData = Nothing
    Set xmlCopyNode = Nothing
    Set xmlAccountRelation = Nothing
    Set xmlAccountGuid = Nothing
    Set xmlAccount = Nothing
    Set xmlAccountList = Nothing
    Set xmlRelData = Nothing
    'BMIDS00405
    Set xmlResults = Nothing
    Set xmlAccountExists = Nothing
    'BMIDS00405 End
     Set xmlShowNotRedeemedNode = Nothing 'BM0492 LDM 15/05/2003
   
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

Private Function ICustomerFinancialDO_GetMortgageAccountDetails(ByVal vxmlTableElement As IXMLDOMElement) As IXMLDOMNode
' header ----------------------------------------------------------------------------------
' description:
'   Get Details of the mortgage account
' pass:
'   vxmlTableElement  xml element containing the request
' return:                xml node containing retrieved data
'------------------------------------------------------------------------------------------
On Error GoTo GetMortgageAccountDetailsVbErr

    Const strFunctionName As String = "ICustomerFinancialDO_GetMortgageAccountDetails"
    
            
    Dim xmlClassDefDoc As FreeThreadedDOMDocument40
    Dim objIClassDef As ICFClassDef
    Dim xmlCopy As IXMLDOMNode
    Dim xmlData As IXMLDOMNode
    Dim xmlMortgageAccount As IXMLDOMNode
    Dim xmlAccountRelation As IXMLDOMNode
    Dim xmlAccountGuid As IXMLDOMNode
    Dim xmlRelData As IXMLDOMNode
    Dim xmlNode As IXMLDOMNode
    
    'BMIDS00510 MDC 30/09/2002
    Dim objAddressBO As IAddressBO
    Dim xmlTempDoc As FreeThreadedDOMDocument40
    Dim xmlTempRequest As IXMLDOMNode
    Dim xmlResponse As IXMLDOMNode
    Dim xmlAddressGuid As IXMLDOMNode
    'BMIDS00510 MDC 30/09/2002 - End
    Dim xmlIndemnityGuid As IXMLDOMNode
    
    Set xmlCopy = vxmlTableElement.cloneNode(True)
    m_objXmlAssist.ChangeNodeName xmlCopy, "MORTGAGEACCOUNT", "MORTGAGEACCOUNTDETAILS"
    
    Set objIClassDef = New CFClassDef
    Set xmlClassDefDoc = objIClassDef.LoadMortgageAccountDetailsDataEx()
    Set xmlData = m_objIDOAssist.FindListMultiple(xmlCopy, xmlClassDefDoc)
    
    m_objXmlAssist.ChangeNodeName xmlData, "MORTGAGEACCOUNTDETAILS", "MORTGAGEACCOUNT"
    Set xmlMortgageAccount = m_objXmlAssist.GetMandatoryNode(xmlData, "MORTGAGEACCOUNT")
    
    ' Remove directory if directory guid is null
    Set xmlNode = m_objXmlAssist.GetNode(xmlMortgageAccount, "NAMEANDADDRESSDIRECTORY[DIRECTORYGUID = '']", False)
    
    If Not xmlNode Is Nothing Then
        xmlNode.parentNode.removeChild xmlNode
    End If
    
    ' Remove third party if third party guid is null
    Set xmlNode = m_objXmlAssist.GetNode(xmlMortgageAccount, "THIRDPARTY[THIRDPARTYGUID = '']", False)
    
    If Not xmlNode Is Nothing Then
        xmlNode.parentNode.removeChild xmlNode
    End If
    '//MAR1448 M Heys 24/04/2006 start
    ' Remove Indemnity Insurance if Indemnity guid is null
    Set xmlNode = m_objXmlAssist.GetNode(xmlMortgageAccount, "INDEMNITYINSURANCE[INDEMNITYACCOUNTGUID = '']", False)
    
    If Not xmlNode Is Nothing Then
        xmlNode.parentNode.removeChild xmlNode
    Else
        Set xmlNode = m_objXmlAssist.GetNode(xmlMortgageAccount, "INDEMNITYINSURANCE/INDEMNITYACCOUNTGUID", False)
        xmlNode.parentNode.removeChild xmlNode
    End If
    '//MAR1448 M Heys 24/04/2006 end
    
    ' Get the relationships
    Set xmlAccountRelation = vxmlTableElement.ownerDocument.createElement("ACCOUNTRELATIONSHIP")
    
    'BMIDS00510 MDC 30/09/2002 - BMIDS00518 can be rolled back as a ClassDef amendment makes it obsolete
    'BMIDS00518 MDC 27/09/2002
    Set xmlAccountGuid = m_objXmlAssist.GetMandatoryNode(xmlMortgageAccount, "ACCOUNTGUID")
    'Set xmlAccountGuid = m_objXmlAssist.GetMandatoryNode(xmlMortgageAccount, "MORTGAGEACCOUNT/ACCOUNTGUID")
    'BMIDS00518 MDC 27/09/2002 - End
    'BMIDS00510 MDC 30/09/2002 - End
    
    xmlAccountRelation.appendChild xmlAccountGuid.cloneNode(True)
    Set xmlRelData = ICustomerFinancialDO_FindAccountRelationshipList(xmlAccountRelation)
    xmlMortgageAccount.appendChild xmlRelData
    
    'BMIDS00510 MDC 30/09/2002 - Get Security Address
    Set xmlAddressGuid = m_objXmlAssist.GetMandatoryNode(xmlMortgageAccount, "MAADDRESSGUID")
    'BMIDS00919 MDC 12/11/2002 - Check AddressGuid is present
    If Len(xmlAddressGuid.Text) > 0 Then
        Set xmlTempDoc = New FreeThreadedDOMDocument40
        xmlTempDoc.validateOnParse = False
        xmlTempDoc.setProperty "NewParser", True
        Set xmlTempRequest = vxmlTableElement.ownerDocument.documentElement.cloneNode(False)
        Set xmlNode = xmlTempDoc.createElement("ADDRESS")
        xmlTempRequest.appendChild xmlNode
        m_objXmlAssist.CopyNode "MAADDRESSGUID", "ADDRESSGUID", xmlMortgageAccount, xmlNode, True
        
        Set objAddressBO = m_objContext.CreateInstance(gstrTHIRDPARTY_COMPONENT & ".AddressBO")
        Set xmlResponse = objAddressBO.GetAddress(xmlTempRequest)
        'm_objErrAssist.CheckXMLResponse xmlResponse.xml, True
        m_objXmlAssist.CopyNode "ADDRESS", "SECURITYADDRESS", xmlResponse, xmlMortgageAccount, True
    End If
    'BMIDS00919 MDC 12/11/2002 - End
    'BMIDS00510 MDC 30/09/2002 - End

    Set ICustomerFinancialDO_GetMortgageAccountDetails = xmlMortgageAccount
    
    Set objIClassDef = Nothing
    Set xmlClassDefDoc = Nothing
    Set xmlCopy = Nothing
    Set xmlData = Nothing
    Set xmlMortgageAccount = Nothing
    Set xmlAccountRelation = Nothing
    Set xmlAccountGuid = Nothing
    Set xmlRelData = Nothing
    Set xmlNode = Nothing
    Set xmlIndemnityGuid = Nothing '//MAR1448 M Heys 24/04/2006

    m_objContext.SetComplete
    
    
    Exit Function

GetMortgageAccountDetailsVbErr:
    
    
    Set objIClassDef = Nothing
    Set xmlClassDefDoc = Nothing
    Set xmlCopy = Nothing
    Set xmlData = Nothing
    Set xmlMortgageAccount = Nothing
    Set xmlAccountRelation = Nothing
    Set xmlAccountGuid = Nothing
    Set xmlRelData = Nothing
    Set xmlNode = Nothing
    Set xmlIndemnityGuid = Nothing '//MAR1448 M Heys 24/04/2006
    
    'BMIDS00510 MDC 30/09/2002
    Set objAddressBO = Nothing
    Set xmlTempDoc = Nothing
    Set xmlTempRequest = Nothing
    Set xmlResponse = Nothing
    Set xmlAddressGuid = Nothing
    'BMIDS00510 MDC 30/09/2002 - End
    
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

Private Function ICustomerFinancialDO_FindMortgageListForArrears(ByVal vxmlTableElement As IXMLDOMElement) As IXMLDOMNode
' header ----------------------------------------------------------------------------------
' description:
'   Get the data for all instances of the persistant data associated with
'   this data object for the values supplied
' pass:
'   vxmlxmlTableElement  xml element containing the request
' return:                xml node containing retrieved data
'------------------------------------------------------------------------------------------
On Error GoTo FindMortgageListForArrearsVbErr

    Const strFunctionName As String = "ICustomerFinancialDO_FindMortgageListForArrears"
    
    Dim xmlClassDefDoc As FreeThreadedDOMDocument40
    Dim objICFClassDef As ICFClassDef
    Dim xmlReturn As IXMLDOMNode
    'BMIDS00190
    Dim xmlCustomers    As IXMLDOMNode
    Dim xmlCustomer     As IXMLDOMNode
    Dim xmlAccount      As IXMLDOMNode
    Dim strAccountGUID  As String
    'BMIDS00190 End
     
    Set objICFClassDef = New CFClassDef
    'SYS0679 Use LoadMortgageAccountThirdPartyData rather than LoadMortgageAccountDetailsData
    'SYS0419 BG 03/10/00 Use LoadMortgageAccountArrearsData rather than LoadMortgageAccountThirdPartyData
    Set xmlClassDefDoc = objICFClassDef.LoadMortgageAccountArrearsData()
    Set xmlCustomers = m_objIDOAssist.FindListMultiple(vxmlTableElement, xmlClassDefDoc, "ACCOUNTGUID, CUSTOMERNUMBER", "MORTGAGEACCOUNT")

    'BMIDS00190
    If Not xmlCustomers Is Nothing Then
        Set xmlReturn = xmlClassDefDoc.createElement("MORTGAGEGACCOUNTARREARSLIST")
        For Each xmlCustomer In xmlCustomers.selectNodes("MORTGAGEACCOUNTARREARS")
            strAccountGUID = xmlCustomer.selectSingleNode("ACCOUNTGUID").Text
            Set xmlAccount = xmlReturn.selectSingleNode("MORTGAGEACCOUNTARREARS[ACCOUNTGUID[.='" & strAccountGUID & "']]")
            If xmlAccount Is Nothing Then
                Set xmlAccount = xmlCustomer
                xmlReturn.appendChild xmlAccount
            Else
                xmlAccount.appendChild xmlCustomer.selectSingleNode("ACCOUNTRELATIONSHIP")
            End If
        Next
    End If
    'BMIDS00190 End

    Set ICustomerFinancialDO_FindMortgageListForArrears = xmlReturn

    Set objICFClassDef = Nothing
    Set xmlClassDefDoc = Nothing
    Set xmlReturn = Nothing
    'BMIDS00190
    Set xmlCustomers = Nothing
    Set xmlCustomer = Nothing
    Set xmlAccount = Nothing
    'BMIDS00190 End
    
    If Not m_objContext Is Nothing Then
        m_objContext.SetComplete
    End If

    Exit Function

FindMortgageListForArrearsVbErr:
    
    Set objICFClassDef = Nothing
    Set xmlClassDefDoc = Nothing
    Set xmlReturn = Nothing
    'BMIDS00190
    Set xmlCustomers = Nothing
    Set xmlCustomer = Nothing
    Set xmlAccount = Nothing
    'BMIDS00190 End

    m_objErrAssist.AddToErrSource strFunctionName
    
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

'BMIDS00190
Private Function ICustomerFinancialDO_FindLoansListForArrears(ByVal vxmlTableElement As IXMLDOMElement) As IXMLDOMNode
' header ----------------------------------------------------------------------------------
' description:
'   Get the data for all instances of the persistant data associated with
'   this data object for the values supplied
' pass:
'   vxmlxmlTableElement  xml element containing the request
' return:                xml node containing retrieved data
'------------------------------------------------------------------------------------------
On Error GoTo FindLoansListForArrearsVbErr

    Const strFunctionName As String = "ICustomerFinancialDO_FindLoansListForArrears"
    
    Dim xmlClassDefDoc  As FreeThreadedDOMDocument40
    Dim objICFClassDef  As ICFClassDef
    Dim xmlReturn       As IXMLDOMNode
    Dim xmlCustomers    As IXMLDOMNode
    Dim xmlCustomer     As IXMLDOMNode
    Dim xmlAccount      As IXMLDOMNode
    Dim strAccountGUID  As String
     
    Set objICFClassDef = New CFClassDef
    Set xmlClassDefDoc = objICFClassDef.LoadLoansLiabilitiesArrearsData()
    Set xmlCustomers = m_objIDOAssist.FindListMultiple(vxmlTableElement, xmlClassDefDoc, "ACCOUNTGUID, CUSTOMERNUMBER", "LOANSLIABILITIES")

    If Not xmlCustomers Is Nothing Then
        Set xmlReturn = xmlClassDefDoc.createElement("LOANSLIABILITIESARREARSLIST")
        For Each xmlCustomer In xmlCustomers.selectNodes("LOANSLIABILITIESARREARS")
            strAccountGUID = xmlCustomer.selectSingleNode("ACCOUNTGUID").Text
            Set xmlAccount = xmlReturn.selectSingleNode("LOANSLIABILITIESARREARS[ACCOUNTGUID[.='" & strAccountGUID & "']]")
            If xmlAccount Is Nothing Then
                Set xmlAccount = xmlCustomer
                xmlReturn.appendChild xmlAccount
            Else
                xmlAccount.appendChild xmlCustomer.selectSingleNode("ACCOUNTRELATIONSHIP")
            End If
        Next
    End If

    Set ICustomerFinancialDO_FindLoansListForArrears = xmlReturn

    Set objICFClassDef = Nothing
    Set xmlClassDefDoc = Nothing
    Set xmlReturn = Nothing
    Set xmlCustomers = Nothing
    Set xmlCustomer = Nothing
    Set xmlAccount = Nothing
    
    If Not m_objContext Is Nothing Then
        m_objContext.SetComplete
    End If

    Exit Function

FindLoansListForArrearsVbErr:
    
    Set objICFClassDef = Nothing
    Set xmlClassDefDoc = Nothing
    Set xmlReturn = Nothing
    Set xmlCustomers = Nothing
    Set xmlCustomer = Nothing
    Set xmlAccount = Nothing

    m_objErrAssist.AddToErrSource strFunctionName
    
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
'BMIDS00190 End

'BMIDS00444
Private Function ICustomerFinancialDO_FindRemortgageAccountAddress(ByVal vxmlTableElement As IXMLDOMElement) As IXMLDOMNode
' header ----------------------------------------------------------------------------------
' description: Get the data for all instances of the persistant data associated with
'              this data object for the values supplied
' pass:     vxmlxmlTableElement     xml element containing the request
' return:   xml node containing retrieved data
'------------------------------------------------------------------------------------------
On Error GoTo FindRemortgageAccountAddressVbErr

    Const strFunctionName As String = "ICustomerFinancialDO_FindRemortgageAccountAddress"
    
    Dim xmlClassDefDoc  As FreeThreadedDOMDocument40
    Dim objICFClassDef  As ICFClassDef
    Dim xmlReturn       As IXMLDOMNode
    Dim xmlCustomers    As IXMLDOMNode
    Dim xmlCustomer     As IXMLDOMNode
    Dim xmlAccount      As IXMLDOMNode
    Dim strAccountGUID  As String
     
    Set objICFClassDef = New CFClassDef
    Set xmlClassDefDoc = objICFClassDef.LoadRemortgageAccountAddressData()
    Set xmlCustomers = m_objIDOAssist.FindListMultiple(vxmlTableElement, xmlClassDefDoc, "ACCOUNTGUID, CUSTOMERNUMBER", "ACCOUNTRELATIONSHIP")

    If Not xmlCustomers Is Nothing Then
        Set xmlReturn = xmlClassDefDoc.createElement("MORTGAGEGACCOUNTLIST")
        For Each xmlCustomer In xmlCustomers.selectNodes("FINDREMORTGAGEACCOUNTADDRESS")
            strAccountGUID = xmlCustomer.selectSingleNode("ACCOUNTGUID").Text
            Set xmlAccount = xmlReturn.selectSingleNode("MORTGAGEACCOUNT[ACCOUNTGUID[.='" & strAccountGUID & "']]")
            If xmlAccount Is Nothing Then
                m_objXmlAssist.ChangeNodeName xmlCustomer, "FINDREMORTGAGEACCOUNTADDRESS", "MORTGAGEACCOUNT"
                Set xmlAccount = xmlCustomer
                xmlReturn.appendChild xmlAccount
            Else
                xmlAccount.appendChild xmlCustomer.selectSingleNode("ACCOUNTRELATIONSHIP")
            End If
        Next
    End If

    Set ICustomerFinancialDO_FindRemortgageAccountAddress = xmlReturn

    Set objICFClassDef = Nothing
    Set xmlClassDefDoc = Nothing
    Set xmlReturn = Nothing
    Set xmlCustomers = Nothing
    Set xmlCustomer = Nothing
    Set xmlAccount = Nothing
    
    If Not m_objContext Is Nothing Then
        m_objContext.SetComplete
    End If

    Exit Function

FindRemortgageAccountAddressVbErr:
    
    Set objICFClassDef = Nothing
    Set xmlClassDefDoc = Nothing
    Set xmlReturn = Nothing
    Set xmlCustomers = Nothing
    Set xmlCustomer = Nothing
    Set xmlAccount = Nothing

    m_objErrAssist.AddToErrSource strFunctionName
    
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
'BMIDS00444 End

'EP2_954 GHun
Private Function ICustomerFinancialDO_GetAdditionalBorrowingDataForProjectionCalc(ByVal vxmlRequest As IXMLDOMElement) As IXMLDOMNode

    Const cstrFunctionName As String = "ICustomerFinancialDO_GetAdditionalBorrowingDataForProjectionCalc"
On Error GoTo GetAdditionalBorrowingDataForProjectionCalcVbErr

    Dim oCommand            As ADODB.Command
    Dim adoConn             As ADODB.Connection
    Dim adoStream           As ADODB.Stream
    Dim objIADOAssist       As IADOAssist
    Dim strAppNum           As String
    Dim strAmountRequested  As String
    Dim strLTV              As String
    Dim strResponse         As String
    Dim xmlDoc              As FreeThreadedDOMDocument40
    Dim xmlInterestRateType As IXMLDOMNode
    Dim xmlNode             As IXMLDOMNode
    Dim xmlChild            As IXMLDOMNode
    Dim xmlBaseRate         As IXMLDOMNode
    
    'Use XML assist and obtain the parameters to pass into the command object.
    strAppNum = m_objXmlAssist.GetNodeValue(vxmlRequest, "APPLICATIONNUMBER")
    strAmountRequested = m_objXmlAssist.GetNodeValue(vxmlRequest, "AMOUNTREQUESTED")
    strLTV = m_objXmlAssist.GetNodeValue(vxmlRequest, "LTV")

    If m_objContext Is Nothing Then
        Set objIADOAssist = New ADOAssist
    Else
        Set objIADOAssist = m_objContext.CreateInstance(gstrBASE_COMPONENT & ".ADOAssist")
    End If

    Set adoConn = New ADODB.Connection
    With adoConn
        .ConnectionString = objIADOAssist.GetConnStr
        .open
    End With
    
    Set adoStream = New ADODB.Stream
    adoStream.open

    'Create a command object to use.
    Set oCommand = New ADODB.Command
    With oCommand
        .CommandType = adCmdStoredProc
        .CommandText = "USP_GetMortgageLoanProductDetails"
        Set .ActiveConnection = adoConn
        .Parameters.Append .CreateParameter("@ApplicationNumber", adVarChar, adParamInput, Len(strAppNum), strAppNum)
        .Parameters.Append .CreateParameter("@AmountRequested", adInteger, adParamInput, , CLng(strAmountRequested))
        .Parameters.Append .CreateParameter("@LTV", adDouble, adParamInput, , CDbl(strLTV))
        .Properties("Output Stream") = adoStream
        .Properties("XML Root") = "MORTGAGELOANLIST"
        .Execute , , adExecuteStream
        Set .ActiveConnection = Nothing
    End With
        
    adoConn.Close

    strResponse = adoStream.ReadText()
    Set xmlDoc = m_objXmlAssist.load(strResponse, TypeName(Me), cstrFunctionName)
    
    'Adjust XML hierarchy to match the XML returned by MortgageProductDO.GetFullProductDetails
    For Each xmlInterestRateType In xmlDoc.selectNodes("MORTGAGELOANLIST/MORTGAGELOAN/INTERESTRATETYPE")
        Set xmlNode = xmlInterestRateType.selectSingleNode("BASERATESET/BASERATEBAND")
        If Not xmlNode Is Nothing Then
            xmlInterestRateType.appendChild xmlNode
        End If
        Set xmlBaseRate = xmlInterestRateType.selectSingleNode("BASERATEBAND")
        Set xmlNode = xmlBaseRate.selectSingleNode("BASERATE")
        For Each xmlChild In xmlNode.childNodes
            xmlBaseRate.appendChild xmlChild
        Next
        xmlBaseRate.removeChild xmlNode
        Set xmlNode = xmlInterestRateType.selectSingleNode("BASERATESET")
        For Each xmlChild In xmlNode.childNodes
            xmlBaseRate.appendChild xmlChild
        Next
        xmlInterestRateType.removeChild xmlNode
        'Rename BASERATEBANDRATE to RATE
        Set xmlChild = xmlDoc.createElement("RATE")
        Set xmlNode = xmlBaseRate.selectSingleNode("BASERATEBANDRATE")
        xmlChild.Text = xmlNode.Text
        xmlBaseRate.appendChild xmlChild
        xmlBaseRate.removeChild xmlNode
    Next
    
    Set ICustomerFinancialDO_GetAdditionalBorrowingDataForProjectionCalc = xmlDoc.documentElement
    
    adoStream.Close
    
    Set adoConn = Nothing
    Set oCommand = Nothing
    Set adoStream = Nothing
    Set objIADOAssist = Nothing
    Set xmlDoc = Nothing
    Set xmlInterestRateType = Nothing
    Set xmlNode = Nothing
    Set xmlBaseRate = Nothing
    Set xmlChild = Nothing

    Exit Function

GetAdditionalBorrowingDataForProjectionCalcVbErr:

    If adoConn.State = adStateOpen Then
       adoConn.Close
    End If
    Set adoConn = Nothing
    Set oCommand = Nothing
    Set adoStream = Nothing
    Set objIADOAssist = Nothing
    Set xmlDoc = Nothing
    Set xmlInterestRateType = Nothing
    Set xmlNode = Nothing
    Set xmlBaseRate = Nothing
    Set xmlChild = Nothing
    
    m_objErrAssist.AddToErrSource (cstrFunctionName)
    
    If m_objErrAssist.IsSystemError = True Then
        m_objErrAssist.LogError TypeName(Me), cstrFunctionName, Err.Number, Err.Description
    End If
    
    If m_objErrAssist.IsSystemError Then
        m_objContext.SetAbort
    Else
        m_objContext.SetComplete
    End If
        
    '   re-raise error for business object to interpret as appropriate
    Err.Raise Err.Number, Err.Source, Err.Description
End Function
'EP2_954 End


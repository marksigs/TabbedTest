VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 3  'UsesTransaction
END
Attribute VB_Name = "MortgageProductDO"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
'Workfile:      MortgageProductDO.cls
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
'Prog   Date        Description
'MCS    07/09/99    Created
'MCS    05/11/99    Type of buyer added for costmodelfindlist 1 & 2
'MCS    19/11/99    LoadMortgage:LenderData modified
'RF     22/11/99    INTERESTRATETYPEGUID removed from LoadInterestRateTypeData
'RF     24/11/99    AQR MC19: FormatInterestRateSQL copes with "LT" to represent "<"
'PSC    02/12/99    Amend to use MortgageProductLanguage
'MCS    03/01/2000  Y2k test
'MCS    05/01/2000  GetValidActiveProduct added with interface
'PSC    28/02/2000  Get default number of records if not entered
'PSC    28/02/2000  Amend LTV parameter to be numeric rather than integer
'APS    09/05/2000  Changed the SQL Select statement on IMortgageProductDO_GetValidActiveProduct
'JLD    12/05/00    SYS0666 - code error in AddInterestRateAndBaseRate error handler.
'JLD    23/05/00    Add output of TypicalApr to FindProductList
'APS    26/05/2000  GetValidActiveProduct SQL select amended to include checks against
'                   Employment status and Member Of Staff indicator
'JLD    30/05/00    Add output of FIRSTMONTHSREPAYMENTBYDAILY, FIRSTMONTHLYSVRREPAYBYDAILY to FindProductList
'PSC    15/06/00    SYS0763 - Return total number of records in FindProductList
'PSC    10/07/00    SYS1193 - Remove RESERVEDPRODUCT from GetValidActiveProduct()
'MS     21/07/00    performance rework.
'MS     03/03/00    SYS1386 use Rateset in calculation
'APS    18/09/00    GetFullProductDetails VB error
'DJP    26/09/00    GetMIGRate - An error causes a vb error. SYS1584
'JLD    28/09/00    mortgageproductstartdate is now dateTime so call formatDateString correctly.
'BG     17/00/00    SYS1600 CM150 Fixed the way GetInclusiveIncentivesForProduct and GetExclusiveIncentivesForProduct
'                   methods instantiated Classdefs.
'LD     07/11/00    Explicity close database connections
'LD     07/11/00    Explicity close recordsets
'LD     11/06/01    SYS2367 SQL Server Port - Length must be specified in calls to CreateParameter
'AS     13/06/01    CC012 SQL Server Port - FindProductList
'DM     02/08/01    SYS2494 Format guid
'MDC    03/09/01    SYS2250 New BaseRate table structure. Amended GetBaseRateBand & FindProductList
'MDC    28/09/01    SYS2250 No longer required to change MortgageproductSelection package name
'AD     08/10/01    Removed the MoveFirst as it's a ForwardOnly cursor. SYS2718
'MDC    06/12/01    SYS2250 - Reinstate backed out BaseRate changes
'MDC    10/12/01    SYS3440 - Rate Change amendments
'DRC    15/02/02    SYS4025 - DRC must be able to get bandrateband 100% LTV
'DRC    08/03/02    SYS4176 - Added Functions GetCurrentRateForBaseRateSet, GetInterestRateType, GetRateDifference
'DRC    05/04/02    SYS4125 - Enhanced No Record found message in GetMIG
'STB    01/05/02    SYS4506 - Incerased parameter sizes for LTV and AmountRequested in FindMortgageProducts.
'MEVA   22/05/02    SYS4696 - Revised SQL query
'DB     28/05/02    SYS4767 - MSMS to Core Integration
'------------------------------------------------------------------------------------------
'BMIDS History:
'
'   Prog    Date        Description
'
'   AW      13/06/02    BM0011      Amended 'FindProdouctList' PRODUCTSGROUP is now passed to
'                               usp_MPSFindProducts as an INT.
'   MDC     14/06/2002  BMIDS00074  CMWP6 BM016 Move BaseRateSet from MortgageProduct to InterestRateType
'                                   Amended AddInterestRateAndBaseRate and GetFullProductDetails
'   AW      25/06/2002  BMIDS00077  Core Integration for 'FindProductList'
'   DPF     11/07/2002  CMWP3       1. Added field to search list in 'FindProductList' - ManualPortedLoanInd
'                                   2. Added new method called GetRedemptionFeeBand()
'                                   3. Created call within GetFullProductDetails() to new method in 2.
'   MV      04/09/2002  BMIDS00353  Modified GetMIGRate()
'   GD      10/09/2002  BMIDS00313 - APWP2 - BM004 Added FindProductConditionsList
'   DPF     01/10/2002  BMIDS00558 - Fix for Advanced Application Searching.
'   SA      08/10/2002  BMIDS00246 reapplied - Changes to FindProductList to send in EmploymentStatusList to stored proc.
'   MV      10/10/2002  BMIDS00590/BMIDS00246 Amended FindProductlist and GetMIGRate
'   SA      14/10/2002  BMIDS00627  Modify GetValidActiveProduct so it doesn't use AvailableForSelfEmployedIndicator.
'   MV      18/10/2002  BMIDS00571  Amended GetMIGRate() - in SQL Statement
'   MO      22/10/2002  BMIDS00666  Amended GetMIGRate() - changed SQL Statement
'   DPF     23/10/2002  CPWP1 - Amended GetInclusiveIncentivesForProducts & GetExclusiveIncentivesForProducts
'   SA      23/10/2002  BMIDS00515 Amended GetValidActiveProduct to return Minimum & Maximu LTV
'   MV      25/10/2002  BMIDS00613  Amended GetMIGRate()
'   SA      29/10/2002  BMIDS00515 GetValidActiveProdut must allow for no LTV sent in.
'   MV      31/10/2002  BMIDS00613  Amended GetMIGRate()
'   MDC     01/11/2002  BMIDS00654 Added GetMortgageProductIncomeMultiplierSet
'   DPF     06/11/2002  BMIDS00770 Amended GetInclusiveIncentivesForProducts & GetExclusiveIncentivesForProducts
'   SA      15/11/2002  BMIDS00952 GetValidActiveProduct amended to no longer add IsSelfEmployed into SQL
'   SR      02/01/2003  BM0209      Modified GetFullProductDetails
'   MDC     02/01/2003  BM0211      Rate Change Fixes
'   MDC     08/01/2003  BM0230      Failed Batch Runs
'   MDC     13/01/2003  BM0245      Base Rate Band selection
'   KRW     29/04/2004  BMIDS737    Removed RedemptionFeeStartDate from SQL
'   MC      10/06/2004  BMIDS763    GetMIGRate,GetValuationRate and GetAdminFeeAmount methods changed
'   JD      21/06/04    BMIDS765    CC076 added GetRentalIncomeRateSet
'   JD      19/07/04    BMIDS749    CC065 BMIDS749 GetBaseRateBand, GetRateForBaseRateSet changed
'   GHun    29/07/2004  BMIDS821    Amended GetValuationFee, GetAdministrationFee, GetMIGRate
'   GHun    20/08/2004  BMIDS856    Minor cleanup and Amended GetFullProductDetails, GetBaseRateBand
'   GHun    02/09/2004  BMIDS815    Amended GetFullProductDetails, AddInterestRateAndBaseRate, GetRedemptionFeeBand
'   GHun    08/09/2004  BMIDS815    Amended GetRedemptionFeeBand
'   GHun    10/09/2004  BMIDS872    Amended GetBaseRateBand
'   GHun    20/09/2004  BMIDS880    Amended GetBaseRateBand
'   GHun    06/10/2004  BMIDS906    Amended GetRedemptionFeeBand to round PeriodDifference up
'   HMA     13/12/2004  BMIDS959    Remove MortgageProductBands table
'-------------------------------------------------------------------------------------------
'MARS History:
'   HMA     18/08/2005  MAR28       Amended GetValuationFee
'   GHun    12/10/2005  MAR46       Changed AddInterestRateAndBaseRate
'   GHun    15/11/2005  MAR355      Changed AddInterestRateAndBaseRate
'   PE      10/02/2006  MAR1190     Add fields to GetBaseRateBand. (RATEDESCRIPTION,BASERATESTARTDATE,BASERATESETDESCRIPTION)
'   PSC     15/02/2006  MAR1276     Changed GetValuationFee to use application type
'   PE      08/05/2006  MAR1702     Maximum Applicant Age Incorrect
'   PE      08/05/2006  MAR1702     Calculate age correctly!!!
'   PE      12/05/2006  MAR1757     Handle date of birth not set.
'---------------------------------------------------------------------------------------------------------------------------
'EPSOM History:
'   IK      17/06/2006  EP776       return TT Fee from TTFEEBAND
'   PE      14/07/2006  EP974       SQL change in addTTFee to allow non-numeric product codes to be selected.
'   GHun    18/10/2006  EP2_9       Merged MAR1687, MAR1878 and made changes for E2CR27, E2CR49 and E2CR52
'   SR      14/11/2006  EP2_76      Modified FindProductsList to add new adoParameters corresponding to ProductExclusivity
'   SR      16/11/2006  EP2_76      Modified FindProductsList
'   GHun    17/11/2006  EP2_126     Changed GetFullProductDetails to make application number optional
'   MAH     24/11/2006  EP2_196     Formatted Date to retrieve rentalIncomerateset correctly
'   PE      19/11/2006  EP2_898     Formatted Date to retrieve rentalIncomerateset correctly
'   PSC     19/02/2007  EP2_1488    Amend GetFullProductDetails to use ORIGINALLTV if passed in
'----------------------------------------------------------------------------------------------------------------------------

Option Explicit
Implements IMortgageProductDO

Implements ObjectControl
 '=============================================
 'Variable Declaration Section
 '=============================================

Private m_objContext As ObjectContext
Private m_objErrAssist As ErrAssist
Private m_objXmlAssist As XMLAssist
Private m_objIDOAssist As IDOAssist

Private Function IMortgageProductDO_FindProductNames(ByVal vxmlTableElement As IXMLDOMElement) As IXMLDOMNode
' header ----------------------------------------------------------------------------------
' description:
'   Get the data for all instances of the persistant data associated with
'   this data object for the values supplied
' pass:
'   vxmlTableElement  xml element containing the request
' return:                xml node containing retrieved data
'------------------------------------------------------------------------------------------
On Error GoTo FindProductNamesVbErr

    Const strFunctionName As String = "IMortgageProductDO_FindProductNames"
    
    Dim xmlClassDefDoc As FreeThreadedDOMDocument40
    Dim xmlOut As FreeThreadedDOMDocument40
    Dim objIClassDef As IomMPClassDef
    Dim objIADOAssist As IADOAssist
    Dim xmlListElem As IXMLDOMElement
    Dim objSqlAssist As SQLAssist
    Dim adoConnection As ADODB.Connection
    Dim adoCommand As ADODB.Command
    Dim adoRecordset As ADODB.Recordset
    Dim adoParameter As ADODB.Parameter
    
    Set objSqlAssist = New SQLAssist
    Set objIADOAssist = New ADOAssist
    
    Set adoConnection = New ADODB.Connection
    Set adoCommand = New ADODB.Command
    Set adoRecordset = New ADODB.Recordset
    Dim strToDay As String
    Dim strSQL As String
    Dim strOrganisationID As String
    Dim strValidProduct As String
    
    strOrganisationID = m_objXmlAssist.GetNodeValue(vxmlTableElement, _
                                                       ".//ORGANISATIONID", False)
    strValidProduct = m_objXmlAssist.GetNodeValue(vxmlTableElement, _
                                                       ".//VALIDPRODUCT", False)
    
    strSQL = "SELECT MORTGAGEPRODUCT.MORTGAGEPRODUCTCODE, " & _
                    "MORTGAGEPRODUCT.ORGANISATIONID, " & _
                    "MORTGAGEPRODUCT.STARTDATE, " & _
                    "MORTGAGEPRODUCTLANGUAGE.PRODUCTNAME " & _
             "FROM MORTGAGEPRODUCT " & _
             "LEFT JOIN MORTGAGEPRODUCTLANGUAGE ON MORTGAGEPRODUCT.MORTGAGEPRODUCTCODE = MORTGAGEPRODUCTLANGUAGE.MORTGAGEPRODUCTCODE " & _
             "AND MORTGAGEPRODUCT.STARTDATE = MORTGAGEPRODUCTLANGUAGE.STARTDATE "
    
    If Len(strValidProduct) = 0 Then
        strSQL = strSQL & " WHERE MORTGAGEPRODUCT.VALIDPRODUCT =  1 "
    Else
        strSQL = strSQL & " WHERE MORTGAGEPRODUCT.VALIDPRODUCT = " & strValidProduct
    End If
    
    If Len(strOrganisationID) > 0 Then
        'DM 02/08/01 SYS2494 Format guid
        strSQL = strSQL & " AND MORTGAGEPRODUCT.ORGANISATIONID = " & objSqlAssist.FormatGuid(strOrganisationID)
    End If
    
    'Store Todays date based on Database
    strToDay = objSqlAssist.GetSystemDate
    
    strSQL = strSQL & " AND MORTGAGEPRODUCT.STARTDATE <= " & strToDay
    strSQL = strSQL & " AND ( MORTGAGEPRODUCT.ENDDATE  >  " & strToDay
    strSQL = strSQL & " OR  MORTGAGEPRODUCT.STARTDATE IS NULL )"
    
    adoCommand.CommandText = strSQL
    
    'Set Recordset properties. Must be adUseClient for disconnected recordset
    adoRecordset.CursorLocation = adUseClient
    adoRecordset.CursorType = adOpenForwardOnly
    adoRecordset.LockType = adLockReadOnly
    
    adoConnection.Open objIADOAssist.GetConnStr
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
     
    Set objIClassDef = New omMPClassDef
    Set xmlClassDefDoc = objIClassDef.LoadProductNamesData
    
    Set xmlOut = New FreeThreadedDOMDocument40
    xmlOut.validateOnParse = False
    xmlOut.setProperty "NewParser", True

    Set xmlListElem = xmlOut.createElement("MORTGAGEPRODUCTLIST")
    xmlOut.appendChild xmlListElem

    ' Convert recordset to XML
    m_objIDOAssist.GetXMLFromWholeRecordset adoRecordset, xmlClassDefDoc, xmlListElem
    
    'Close recordset Object
    adoRecordset.Close
    
    Set IMortgageProductDO_FindProductNames = xmlListElem
    
    Set objIClassDef = Nothing
    Set xmlClassDefDoc = Nothing
    Set xmlOut = Nothing
    Set objIADOAssist = Nothing
    Set xmlListElem = Nothing
    Set adoConnection = Nothing
    Set adoCommand = Nothing
    Set adoRecordset = Nothing
    Set adoParameter = Nothing
    Set objSqlAssist = Nothing
    
    m_objContext.SetComplete
    
    Exit Function

FindProductNamesVbErr:
    
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
    Set xmlClassDefDoc = Nothing
    Set xmlOut = Nothing
    Set objIADOAssist = Nothing
    Set xmlListElem = Nothing
    Set adoConnection = Nothing
    Set adoCommand = Nothing
    Set adoRecordset = Nothing
    Set adoParameter = Nothing
    Set objSqlAssist = Nothing
    
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

'BMIDS00654 MDC 01/11/2002
Private Function IMortgageProductDO_GetMortgageProductIncomeMultiplierSet(ByVal vxmlTableElement As MSXML2.IXMLDOMElement) As MSXML2.IXMLDOMNode

On Error GoTo GetMortgageProductIncomeMultiplierSetVbErr

    Const strFunctionName As String = "IMortgageProductDO_GetMortgageProductIncomeMultiplierSet"
    
    Dim xmlClassDefDoc As FreeThreadedDOMDocument40
    Dim objIomMpClassDef As IomMPClassDef
                    
    Set objIomMpClassDef = New omMPClassDef
    Set xmlClassDefDoc = objIomMpClassDef.LoadMortgageProductIncomeMultiplierSet
    Set IMortgageProductDO_GetMortgageProductIncomeMultiplierSet = m_objIDOAssist.GetData(vxmlTableElement, xmlClassDefDoc, "MORTGAGEPRODUCT")

    Set objIomMpClassDef = Nothing
    Set xmlClassDefDoc = Nothing
    
    If Not m_objContext Is Nothing Then
        m_objContext.SetComplete
    End If
    
    Exit Function

GetMortgageProductIncomeMultiplierSetVbErr:
    
    Set objIomMpClassDef = Nothing
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

End Function
'BMIDS00654 MDC 01/11/2002 - End

Private Function IMortgageProductDO_GetMortgageProductRentalIncomeRateSet(ByVal vxmlTableElement As IXMLDOMElement) As IXMLDOMNode
' header ----------------------------------------------------------------------------------
' description:
'   Get associated data from the rentalIncomeRateSetBand table
' pass:
'   vxmlxmlTableElement  xml element containing the request
' return:                xml node containing retrieved data
'------------------------------------------------------------------------------------------
On Error GoTo GetMortgageProductRentalIncomeVbErr

Const strFunctionName As String = "IMortgageProductDO_GetMortgageProductRentalIncomeRateSet"
    
Dim xmlClassDefDoc As FreeThreadedDOMDocument40
Dim objIomMpClassDef As IomMPClassDef
Dim objIADOAssist As IADOAssist
Dim xmlNode As IXMLDOMNode
Dim xmlTopNode As IXMLDOMNode
Dim objComm As ADODB.Command
Dim objRs As ADODB.Recordset

Dim strSQL As String
Dim strRateSet As String
Dim strStartDate As String
Dim strMaxStartDate As String
Dim strLoanAmount As String
Dim strLTV As String
           
    If m_objContext Is Nothing Then
        Set objIADOAssist = New ADOAssist
    Else
        Set objIADOAssist = m_objContext.CreateInstance(gstrBASE_COMPONENT & ".ADOAssist")
    End If
    
    Set objComm = New ADODB.Command
    
    Set xmlTopNode = m_objXmlAssist.GetMandatoryNode(vxmlTableElement, "RENTALINCOMERATESETBAND")
    Set xmlNode = m_objXmlAssist.GetMandatoryNode(xmlTopNode, "RENTALINCOMERATESET")
    strRateSet = xmlNode.Text
    Set xmlNode = m_objXmlAssist.GetMandatoryNode(xmlTopNode, "STARTDATE")
    strStartDate = xmlNode.Text
    Set xmlNode = m_objXmlAssist.GetMandatoryNode(xmlTopNode, "LOANAMOUNT")
    strLoanAmount = xmlNode.Text
    Set xmlNode = m_objXmlAssist.GetMandatoryNode(xmlTopNode, "LTV")
    strLTV = xmlNode.Text
    
    ' Find the rental income interest rate
    strMaxStartDate = "(SELECT MAX(STARTDATE) FROM RENTALINCOMERATESETBAND WHERE "
    strMaxStartDate = strMaxStartDate & "RENTALINCOMERATESET=" & strRateSet & " AND STARTDATE <= '" & Format(strStartDate, "dd-mm-yyyy hh:nn:ss") & "')" 'MAH 24/11/2006 EP2_196
    
    strSQL = "SELECT RB.* FROM RENTALINCOMERATESETBAND RB"
    strSQL = strSQL & " WHERE RB.RENTALINCOMERATESET =" & strRateSet & " AND "
    strSQL = strSQL & "RB.STARTDATE=" & strMaxStartDate & " AND "
    strSQL = strSQL & "RB.MAXLOANAMOUNT=(SELECT MIN(MAXLOANAMOUNT) FROM RENTALINCOMERATESETBAND WHERE "
    strSQL = strSQL & "RENTALINCOMERATESET=" & strRateSet & " AND STARTDATE=" & strMaxStartDate & " AND "
    strSQL = strSQL & "MAXLOANAMOUNT >= " & strLoanAmount & ") AND "
    strSQL = strSQL & "RB.MAXLTV=(SELECT MIN(MAXLTV) FROM RENTALINCOMERATESETBAND WHERE "
    strSQL = strSQL & "RENTALINCOMERATESET=" & strRateSet & " AND STARTDATE=" & strMaxStartDate & " AND "
    strSQL = strSQL & "MAXLTV >= " & strLTV & ")"

    With objComm
        .ActiveConnection = objIADOAssist.GetConnStr
        .CommandType = adCmdText
        .CommandText = strSQL
        Set objRs = .Execute()
    End With
    
    Set objIomMpClassDef = New omMPClassDef
    Set xmlClassDefDoc = objIomMpClassDef.LoadRentalIncomeRateData
    
    If Not objRs.EOF Then
        Set IMortgageProductDO_GetMortgageProductRentalIncomeRateSet = m_objIDOAssist.GetXMLFromRecordSet(objRs, xmlClassDefDoc)
    Else
       ' return null
        Set IMortgageProductDO_GetMortgageProductRentalIncomeRateSet = Nothing
    End If
    
    objRs.Close
    
    Set objIomMpClassDef = Nothing
    Set xmlClassDefDoc = Nothing
    Set objIADOAssist = Nothing
    Set objComm = Nothing
    Set objRs = Nothing
    Set xmlNode = Nothing
    
    If Not m_objContext Is Nothing Then
        m_objContext.SetComplete
    End If
    
    Exit Function

GetMortgageProductRentalIncomeVbErr:
    
    Set objIomMpClassDef = Nothing
    Set xmlClassDefDoc = Nothing
    Set xmlNode = Nothing
    
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

Private Function IMortgageProductDO_GetRateForBaseRateSet(ByVal vxmlTableElement As IXMLDOMElement) As IXMLDOMNode
' header ----------------------------------------------------------------------------------
' description:
'   Get associated data from the BaseRate table for a specified BaseRateSet
' pass:
'   vxmlxmlTableElement  xml element containing the request
' return:                xml node containing retrieved data
'------------------------------------------------------------------------------------------
On Error GoTo GetRateForBaseRateSetVbErr

Const strFunctionName As String = "IMortgageProductDO_GetRateForBaseRateSet"
    
Dim xmlClassDefDoc As FreeThreadedDOMDocument40
Dim objIomMpClassDef As IomMPClassDef
Dim objIADOAssist As IADOAssist
Dim objSqlAssist As SQLAssist

'SYS3440 MDC 10/12/2001
Dim intDays As Integer
Dim blnRateChange As Boolean
Dim strValue As String
Dim xmlNode As IXMLDOMNode
Dim objGlobalParameterDO As GlobalParameterDO
'SYS3440 End

Dim objComm As ADODB.Command
Dim objRs As ADODB.Recordset

Dim strSQL As String
Dim strBaseRateSet As String
            
    If m_objContext Is Nothing Then
        Set objIADOAssist = New ADOAssist
        Set objGlobalParameterDO = New GlobalParameterDO    'SYS3440 MDC 10/12/2001
    Else
        Set objIADOAssist = m_objContext.CreateInstance(gstrBASE_COMPONENT & ".ADOAssist")
        Set objGlobalParameterDO = m_objContext.CreateInstance(gstrBASE_COMPONENT & ".GlobalParameterDO") 'SYS3440 MDC 10/12/2001
    End If
    
    Set objComm = New ADODB.Command
    
    'SYS3440 MDC 10/12/2001
    Set xmlNode = m_objXmlAssist.GetMandatoryNode(vxmlTableElement, "BASERATESET")
    'Get the Base Rate Set
    'strBaseRateSet = vxmlTableElement.Text
    strBaseRateSet = xmlNode.Text
    Set xmlNode = vxmlTableElement.selectSingleNode("RATECHANGE")
    If Not xmlNode Is Nothing Then
        blnRateChange = CBool(xmlNode.Text)
    End If
    If blnRateChange Then
        strValue = objGlobalParameterDO.GetCurrentParameterByType("RateChangeLookAheadDays", "AMOUNT")
        intDays = CInt(strValue)
    End If
    'SYS3440 End
    
    Set objSqlAssist = New SQLAssist
    
    ' Find the current Base Rate record for the BaseRateSet
    strSQL = "SELECT BR.* FROM BASERATE BR, BASERATESET BS"
    strSQL = strSQL & " WHERE BR.RATEID = BS.RATEID AND BR.BASERATESTARTDATE ="
    strSQL = strSQL & " (SELECT MAX(BASERATESTARTDATE) FROM BASERATE"
    strSQL = strSQL & " WHERE RATEID = BR.RATEID AND BASERATESTARTDATE <= "
    strSQL = strSQL & objSqlAssist.GetSystemDate()
    'SYS3440 MDC 10/12/2001
    If intDays > 0 Then
        strSQL = strSQL & " + " & CStr(intDays)
    End If
    strSQL = strSQL & ")"
    'SYS3440 End
    strSQL = strSQL & " AND BS.BASERATESET = '" & strBaseRateSet & "'"
    ' JD BMIDS749
    If blnRateChange Then
        strSQL = strSQL & " AND BASERATEAPPLIEDDATE IS NULL"
    End If

    With objComm
        .ActiveConnection = objIADOAssist.GetConnStr
        .CommandType = adCmdText
        .CommandText = strSQL
        Set objRs = .Execute()
    End With
    
    Set objIomMpClassDef = New omMPClassDef
    Set xmlClassDefDoc = objIomMpClassDef.LoadBaseRateData
    
    If Not objRs.EOF Then
        Set IMortgageProductDO_GetRateForBaseRateSet = m_objIDOAssist.GetXMLFromRecordSet(objRs, xmlClassDefDoc)
    Else
        ' raise application error to be interpreted by calling object
        m_objErrAssist.RaiseError TypeName(Me), strFunctionName, omiga4RecordNotFound, "BaseRate"
    End If
    
    objRs.Close
    
    Set objSqlAssist = Nothing
    Set objIomMpClassDef = Nothing
    Set xmlClassDefDoc = Nothing
    Set objIADOAssist = Nothing
    Set objComm = Nothing
    Set objRs = Nothing
    'SYS3440 MDC 10/12/2001
    Set xmlNode = Nothing
    Set objGlobalParameterDO = Nothing
    'SYS3440 End
    
    If Not m_objContext Is Nothing Then
        m_objContext.SetComplete
    End If
    
    Exit Function

GetRateForBaseRateSetVbErr:
    
    Set objIomMpClassDef = Nothing
    Set xmlClassDefDoc = Nothing
    
    'SYS3440 MDC 10/12/2001
    Set xmlNode = Nothing
    Set objGlobalParameterDO = Nothing
    'SYS3440 End
    
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

Private Function IMortgageProductDO_GetCurrentRateForBaseRateSet(ByVal vxmlTableElement As IXMLDOMElement) As IXMLDOMNode
' header ----------------------------------------------------------------------------------
' description:
'   Get current rate from the BaseRate table for a specified BaseRateSet
' pass:
'   vxmlxmlTableElement  xml element containing the request
' return:                xml node containing retrieved data
'------------------------------------------------------------------------------------------
On Error GoTo GetCurrentRateForBaseRateSetVbErr

Const strFunctionName As String = "IMortgageProductDO_GetCurrentRateForBaseRateSet"
    
Dim xmlClassDefDoc As FreeThreadedDOMDocument40
Dim objIomMpClassDef As IomMPClassDef
Dim objIADOAssist As IADOAssist

Dim xmlNode As IXMLDOMNode

Dim objComm As ADODB.Command
Dim objRs As ADODB.Recordset

Dim strSQL As String
Dim strBaseRateSet As String
            
    If m_objContext Is Nothing Then
        Set objIADOAssist = New ADOAssist
    Else
        Set objIADOAssist = m_objContext.CreateInstance(gstrBASE_COMPONENT & ".ADOAssist")
    End If
    
    Set objComm = New ADODB.Command
    Set xmlNode = m_objXmlAssist.GetMandatoryNode(vxmlTableElement, "BASERATESET")
    'Get the Base Rate Set
    strBaseRateSet = xmlNode.Text
    
    ' Find the current Base Rate record for the BaseRateSet and the applied
    ' date has not yet been set
    
    strSQL = "SELECT BR.* FROM BASERATE BR, BASERATESET BS"
    strSQL = strSQL & " WHERE BR.RATEID = BS.RATEID AND BR.BASERATESTARTDATE ="
    strSQL = strSQL & " (SELECT MAX(BASERATESTARTDATE) FROM BASERATE"
    strSQL = strSQL & " WHERE RATEID = BR.RATEID AND BASERATEAPPLIEDDATE IS NOT NULL)"
    strSQL = strSQL & " AND BS.BASERATESET = '" & strBaseRateSet & "'"

    With objComm
        .ActiveConnection = objIADOAssist.GetConnStr
        .CommandType = adCmdText
        .CommandText = strSQL
        Set objRs = .Execute()
    End With
    
    Set objIomMpClassDef = New omMPClassDef
    Set xmlClassDefDoc = objIomMpClassDef.LoadBaseRateData
    
    If Not objRs.EOF Then
     Set IMortgageProductDO_GetCurrentRateForBaseRateSet = m_objIDOAssist.GetXMLFromRecordSet(objRs, xmlClassDefDoc)
 
    Else
        ' raise application error to be interpreted by calling object
        m_objErrAssist.RaiseError TypeName(Me), strFunctionName, omiga4RecordNotFound, "BaseRate"
    End If
    
    objRs.Close
    
    Set objIomMpClassDef = Nothing
    Set xmlClassDefDoc = Nothing
    Set objIADOAssist = Nothing
    Set objComm = Nothing
    Set objRs = Nothing
    Set xmlNode = Nothing
    
    If Not m_objContext Is Nothing Then
        m_objContext.SetComplete
    End If
    
    Exit Function

GetCurrentRateForBaseRateSetVbErr:
    
    Set objIomMpClassDef = Nothing
    Set xmlClassDefDoc = Nothing
    Set xmlNode = Nothing
    
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

Private Function IMortgageProductDO_GetRateDifference(ByVal vxmlTableElement As IXMLDOMElement) As IXMLDOMNode
' header ----------------------------------------------------------------------------------
' description:
'   Get the Rate Difference  from the BaseRateBand table for a specified BaseRateSet
' pass:
'   vxmlxmlTableElement  xml element containing the request
' return:                xml node containing retrieved data
'------------------------------------------------------------------------------------------
On Error GoTo GetRateDifferenceVbErr

Const strFunctionName As String = "IMortgageProductDO_GetRateDifference"
    
Dim xmlClassDefDoc As FreeThreadedDOMDocument40
Dim objIomMpClassDef As IomMPClassDef
Dim objIADOAssist As IADOAssist
Dim objSqlAssist As SQLAssist
Dim xmlNode As IXMLDOMNode
Dim objXmlOut As FreeThreadedDOMDocument40
Dim objXmlElem As IXMLDOMElement

Dim objComm As ADODB.Command
Dim objRs As ADODB.Recordset

Dim strSQL As String
Dim strBaseRateSet As String
Dim strLoanAmount As String
Dim strLTV As String
            
  If m_objContext Is Nothing Then
        Set objIADOAssist = New ADOAssist
    Else
        Set objIADOAssist = m_objContext.CreateInstance(gstrBASE_COMPONENT & ".ADOAssist")
    End If
    
    Set objComm = New ADODB.Command
    
    Set xmlNode = m_objXmlAssist.GetMandatoryNode(vxmlTableElement, "BASERATESET")
    'Get the Base Rate Set
    strBaseRateSet = xmlNode.Text
    Set xmlNode = m_objXmlAssist.GetMandatoryNode(vxmlTableElement, "LOANAMOUNT")
    strLoanAmount = xmlNode.Text
    Set xmlNode = m_objXmlAssist.GetMandatoryNode(vxmlTableElement, "LTV")
    strLTV = xmlNode.Text
    
    Set objSqlAssist = New SQLAssist
    
    ' now get the ratedifference (if any)
strSQL = "SELECT * FROM "
strSQL = strSQL & " (SELECT BAND1.BASERATESET, BAND1.BASERATEBANDSTARTDATE,"
strSQL = strSQL & "   BAND1.MAXIMUMLOANAMOUNT, BAND1.MAXIMUMLTV,"
strSQL = strSQL & "   BAND1.RATEDIFFERENCE"
strSQL = strSQL & "   FROM BASERATEBAND BAND1"
strSQL = strSQL & "   WHERE BAND1.BASERATESET = " & strBaseRateSet
strSQL = strSQL & "   AND BAND1.BASERATEBANDSTARTDATE ="
strSQL = strSQL & "     (SELECT MAX(BASERATEBANDSTARTDATE) FROM BASERATEBAND"
strSQL = strSQL & "      WHERE BASERATESET = BAND1.BASERATESET"
strSQL = strSQL & "      AND BASERATEBANDSTARTDATE < " & objSqlAssist.GetSystemDate & ")"
strSQL = strSQL & "   AND BAND1.MAXIMUMLOANAMOUNT ="
strSQL = strSQL & "     (SELECT MIN(MAXIMUMLOANAMOUNT)"
strSQL = strSQL & "     FROM BASERATEBAND WHERE BASERATESET = " & strBaseRateSet
strSQL = strSQL & "     AND MAXIMUMLOANAMOUNT >= " & strLoanAmount
strSQL = strSQL & "     AND BASERATEBANDSTARTDATE = BAND1.BASERATEBANDSTARTDATE)"
strSQL = strSQL & "  )  Q1"
strSQL = strSQL & "  WHERE MAXIMUMLTV ="
strSQL = strSQL & "    (SELECT MIN(MAXIMUMLTV) FROM"
strSQL = strSQL & "         (SELECT BAND1.MAXIMUMLTV FROM BASERATEBAND BAND1"
strSQL = strSQL & "             WHERE  BASERATEBANDSTARTDATE = (SELECT MAX(BASERATEBANDSTARTDATE)"
strSQL = strSQL & "                                             FROM BASERATEBAND"
strSQL = strSQL & "                                              WHERE BASERATESET = " & strBaseRateSet
strSQL = strSQL & "                                              AND BASERATEBANDSTARTDATE < " & objSqlAssist.GetSystemDate & ")"
strSQL = strSQL & "             AND MAXIMUMLOANAMOUNT = (SELECT MIN(MAXIMUMLOANAMOUNT)"
strSQL = strSQL & "                                      FROM BASERATEBAND WHERE BASERATESET = " & strBaseRateSet
strSQL = strSQL & "                                                        AND MAXIMUMLOANAMOUNT >= " & strLoanAmount
strSQL = strSQL & "                                                        AND BASERATEBANDSTARTDATE = BAND1.BASERATEBANDSTARTDATE"
strSQL = strSQL & "                                     )"
strSQL = strSQL & "             AND BAND1.MAXIMUMLOANAMOUNT = (SELECT MIN(MAXIMUMLOANAMOUNT)"
strSQL = strSQL & "                                              FROM BASERATEBAND WHERE BASERATESET = " & strBaseRateSet
strSQL = strSQL & "                                                                AND MAXIMUMLOANAMOUNT >= " & strLoanAmount
strSQL = strSQL & "                                                                AND BASERATEBANDSTARTDATE = BAND1.BASERATEBANDSTARTDATE)"
strSQL = strSQL & "             AND BAND1.MAXIMUMLTV >= " & strLTV
strSQL = strSQL & "          ) Q2"
strSQL = strSQL & "    )"
 
   With objComm
        .ActiveConnection = objIADOAssist.GetConnStr
        .CommandType = adCmdText
        .CommandText = strSQL
        Set objRs = .Execute()
    End With
    
    Set objIomMpClassDef = New omMPClassDef
    Set xmlClassDefDoc = objIomMpClassDef.LoadBaseRateBandData()
    
    If Not objRs.EOF Then
        Set xmlNode = m_objIDOAssist.GetXMLFromRecordSet(objRs, xmlClassDefDoc)
    Else
        Set objXmlOut = New FreeThreadedDOMDocument40
        objXmlOut.validateOnParse = False
        objXmlOut.setProperty "NewParser", True
        Set objXmlElem = objXmlOut.createElement("BASERATEBAND")
        Set xmlNode = objXmlOut.appendChild(objXmlElem)
        Set objXmlElem = objXmlOut.createElement("RATEDIFFERENCE")
        Set xmlNode = xmlNode.appendChild(objXmlElem)
        xmlNode.Text = "0.0"
        Set xmlNode = xmlNode.parentNode
    End If
    Set IMortgageProductDO_GetRateDifference = xmlNode
    objRs.Close
    
    Set objSqlAssist = Nothing
    Set objIomMpClassDef = Nothing
    Set xmlClassDefDoc = Nothing
    Set objIADOAssist = Nothing
    Set objComm = Nothing
    Set objRs = Nothing
    Set xmlNode = Nothing
    
    If Not m_objContext Is Nothing Then
        m_objContext.SetComplete
    End If
    
    Exit Function

GetRateDifferenceVbErr:
   
    Set xmlNode = Nothing
    
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

Private Function IMortgageProductDO_GetData(ByVal vxmlTableElement As IXMLDOMElement) As IXMLDOMNode
' header ----------------------------------------------------------------------------------
' description:
'   Get the data for a single instance of the persistant data associated with
'   this data object
' pass:
'   vxmlxmlTableElement  xml element containing the request
' return:                xml node containing retrieved data
'------------------------------------------------------------------------------------------
On Error GoTo GetDataVbErr

    Const strFunctionName As String = "IMortgageProductDO_GetData"
    
    Dim xmlClassDefDoc As FreeThreadedDOMDocument40
    Dim objIomMpClassDef As IomMPClassDef
                    
    Set objIomMpClassDef = New omMPClassDef
    Set xmlClassDefDoc = objIomMpClassDef.LoadMortgageProductData
    Set IMortgageProductDO_GetData = m_objIDOAssist.GetData(vxmlTableElement, xmlClassDefDoc)

    ' TODO ************************************************************************
    '   Add Any derived data here
    ' *****************************************************************************
    
    Set objIomMpClassDef = Nothing
    Set xmlClassDefDoc = Nothing
    
    If Not m_objContext Is Nothing Then
        m_objContext.SetComplete
    End If
    
    Exit Function

GetDataVbErr:
    
    Set objIomMpClassDef = Nothing
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
End Function

Public Function GetExclusiveIncentivesForProduct(ByVal vstrXMLRequest As String) As String
' header ----------------------------------------------------------------------------------
' description:
'   Get the data for a single instance of the persistant data associated with
'   this data object
' pass:
'   vstrXMLRequest  xml Request data stream containing data to which identifies
'                   the instance of the persistant data to be retrieved
' return:
'   GetData         string containing XML data stream representation of
'                   data retrieved
' Raise Errors: if record not found, raise omiga4RecordNotFound
'------------------------------------------------------------------------------------------
On Error GoTo GetExclusiveIncentivesForProductVbErr

    Const strFunctionName As String = "GetExclusiveIncentivesForProduct"
    
    Dim objSqlAssist As SQLAssist
    Set objSqlAssist = New SQLAssist
                
    Dim rstThisRecSet As ADODB.Recordset
    Dim cmd As ADODB.Command
    Set cmd = New ADODB.Command
    
    Dim objIADOAssist As IADOAssist
    
    Dim objXmlIn As FreeThreadedDOMDocument40
    Dim strSQL As String
         
    Dim strMortgageProductCode As String
    Dim strStartDate As String
    Dim strLoanAmount As String
         
    Set objXmlIn = m_objXmlAssist.load(vstrXMLRequest, TypeName(Me), strFunctionName)
                  
    strMortgageProductCode = m_objXmlAssist.GetTagValue(objXmlIn.documentElement, _
                                         "MORTGAGEPRODUCTCODE")
    strStartDate = m_objXmlAssist.GetTagValue(objXmlIn.documentElement, _
                                                  "STARTDATE")
    strLoanAmount = m_objXmlAssist.GetTagValue(objXmlIn.documentElement, _
                                                  "LOANAMOUNT")

    If Len(strMortgageProductCode) = 0 Or _
       Len(strStartDate) = 0 Then
       Dim strDetails As String
       strDetails = "MORTGAGEPRODUCTCODE = " & strMortgageProductCode
       strDetails = strDetails & "; STARTDATE = " & strStartDate
      
       m_objErrAssist.RaiseError TypeName(Me), _
                               strFunctionName, _
                               omiga4InvalidParameter, _
                               strDetails
    End If
         
    If m_objContext Is Nothing Then
        Set objIADOAssist = New ADOAssist
    Else
        Set objIADOAssist = m_objContext.CreateInstance(gstrBASE_COMPONENT & ".ADOAssist")
    End If
          
       ' Build up the SQL string
    strSQL = "SELECT * FROM INCENTIVE,EXCLUSIVEINCENTIVE WHERE"
    strSQL = strSQL & " MORTGAGEPRODUCTCODE = " & objSqlAssist.FormatString(strMortgageProductCode)
    strSQL = strSQL & " AND STARTDATE = " & objSqlAssist.FormatDate(strStartDate)
    strSQL = strSQL & " AND INCENTIVE.INCENTIVEGUID = EXCLUSIVEINCENTIVE.INCENTIVEGUID"

    cmd.ActiveConnection = objIADOAssist.GetConnStr

    cmd.CommandText = strSQL
    Dim objXmlOldNode As IXMLDOMNode
    
    Set rstThisRecSet = cmd.Execute(1, , adCmdText)
    
    If rstThisRecSet.EOF Then
        ' raise application error to be interpreted by calling object
        m_objErrAssist.RaiseError TypeName(Me), strFunctionName, omiga4RecordNotFound
    End If
    
    Dim objXmlElem As IXMLDOMElement
    'DPF 21/10/02 - CPWP1 add new field
    Dim objXmlIndicatorElem As IXMLDOMElement
    Dim objXmlListNode As IXMLDOMNode
    Dim objXmlOut As FreeThreadedDOMDocument40
    
    Dim objTempElement As IXMLDOMElement
    
    Dim dblPercentage As Double
    Dim dblAmount As Double
    Dim dblLoanAmount As Double
    Dim dblPercentageMaximum As Double
    Dim objConvertAssist As ConvertAssist
    
    Set objXmlOut = New FreeThreadedDOMDocument40
    objXmlOut.validateOnParse = False
    objXmlOut.setProperty "NewParser", True
    Set objXmlElem = objXmlOut.createElement("EXCLUSIVEINCENTIVELIST")
    Set objXmlListNode = objXmlOut.appendChild(objXmlElem)
    
    'BG 17/10/00 SYS1600
    Dim objIomMpClassDef As IomMPClassDef
    Set objIomMpClassDef = New omMPClassDef
    
    Dim xmlLoadIncentiveData As FreeThreadedDOMDocument40
    Set xmlLoadIncentiveData = objIomMpClassDef.LoadIncentiveData
    'BG 17/10/00 End SYS1600
    
    Set objConvertAssist = New ConvertAssist
    
    ' AD 08/10/01. Removed the MoveFirst as it's a ForwardOnly cursor. SYS2718
       ' loop through the record set
    While Not rstThisRecSet.EOF
    
        Set objXmlOldNode = m_objIDOAssist.GetXMLFromRecordSet(rstThisRecSet, xmlLoadIncentiveData)
        
        dblPercentage = objConvertAssist.CSafeDbl(m_objXmlAssist.GetTagValue(objXmlOldNode, "PERCENTAGE"))
        dblLoanAmount = objConvertAssist.CSafeDbl(strLoanAmount)
        
        'DPF 21/10/02 - CPWP1 add new field
        Set objXmlIndicatorElem = objXmlOut.createElement("INCENTIVEFIXEDINDICATOR")
        
        If dblPercentage > 0 Then
            Set objTempElement = objXmlOldNode.ownerDocument.getElementsByTagName("AMOUNT").Item(0)
            dblAmount = ((dblPercentage / 100) * dblLoanAmount)
            dblPercentageMaximum = objConvertAssist.CSafeDbl( _
                            m_objXmlAssist.GetTagValue(objXmlOldNode, "PERCENTAGEMAXIMUM"))
            
            objTempElement.Text = CStr(dblAmount)
            
            If dblPercentageMaximum <> 0 And dblAmount > dblPercentageMaximum Then
                objTempElement.Text = CStr(dblPercentageMaximum)
            End If
            
            'DPF 21/10/02 - CPWP1 add new field
            objXmlIndicatorElem.Text = "0"
            objXmlOldNode.appendChild objXmlIndicatorElem
        Else
            'DPF 21/10/02 - CPWP1 add new field
            objXmlIndicatorElem.Text = "1"
            objXmlOldNode.appendChild objXmlIndicatorElem
        End If

        objXmlListNode.appendChild objXmlOldNode
        
        rstThisRecSet.MoveNext
    Wend
    rstThisRecSet.Close
    
    GetExclusiveIncentivesForProduct = objXmlOut.xml

    Set objSqlAssist = Nothing
    Set rstThisRecSet = Nothing
    Set cmd = Nothing
    Set objIADOAssist = Nothing
    Set objXmlIn = Nothing
    
    If Not m_objContext Is Nothing Then
        m_objContext.SetComplete
    End If
    
    Exit Function

GetExclusiveIncentivesForProductVbErr:
        
    ' Close the recordset if still open
    If Not rstThisRecSet Is Nothing Then
        If rstThisRecSet.State = adStateOpen Then
            rstThisRecSet.Close
        End If
    End If

    Set rstThisRecSet = Nothing
    Set cmd = Nothing
    
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

Public Function GetInclusiveIncentivesForProduct(ByVal vstrXMLRequest As String) As String
' header ----------------------------------------------------------------------------------
' description:
'   Get the data for a single instance of the persistant data associated with
'   this data object
' pass:
'   vstrXMLRequest  xml Request data stream containing data to which identifies
'                   the instance of the persistant data to be retrieved
' return:
'   GetData         string containing XML data stream representation of
'                   data retrieved
' Raise Errors: if record not found, raise omiga4RecordNotFound
'------------------------------------------------------------------------------------------
On Error GoTo GetInclusiveIncentivesForProductVbErr

    Const strFunctionName As String = "GetInclusiveIncentivesForProduct"
     
    Dim objSqlAssist As SQLAssist
    Set objSqlAssist = New SQLAssist
            
    Dim rstThisRecSet As ADODB.Recordset
    Dim cmd As ADODB.Command
    
    Set rstThisRecSet = New ADODB.Recordset
    Set cmd = New ADODB.Command

    Dim objIADOAssist As IADOAssist
    
    Dim objXmlIn As FreeThreadedDOMDocument40
    Dim strSQL As String
         
    Dim strMortgageProductCode As String
    Dim strStartDate As String
    Dim strLoanAmount As String
         
    Set objXmlIn = m_objXmlAssist.load(vstrXMLRequest, TypeName(Me), strFunctionName)
                  
    strMortgageProductCode = m_objXmlAssist.GetTagValue(objXmlIn.documentElement, _
                                         "MORTGAGEPRODUCTCODE")
    strStartDate = m_objXmlAssist.GetTagValue(objXmlIn.documentElement, _
                                                  "STARTDATE")
    strLoanAmount = m_objXmlAssist.GetTagValue(objXmlIn.documentElement, _
                                                  "LOANAMOUNT")
    
    If Len(strMortgageProductCode) = 0 Or _
       Len(strStartDate) = 0 Then
       Dim strDetails As String
       strDetails = "MORTGAGEPRODUCTCODE = " & strMortgageProductCode
       strDetails = strDetails & "; STARTDATE = " & strStartDate
       strDetails = strDetails & "; LOANAMOUNT = " & strLoanAmount
      
       m_objErrAssist.RaiseError TypeName(Me), _
                               strFunctionName, _
                               omiga4InvalidParameter, _
                               strDetails
    End If
         
    If m_objContext Is Nothing Then
        Set objIADOAssist = New ADOAssist
    Else
        Set objIADOAssist = m_objContext.CreateInstance(gstrBASE_COMPONENT & ".ADOAssist")
    End If
          
       ' Build up the SQL string
    strSQL = "SELECT * FROM INCENTIVE,INCLUSIVEINCENTIVE WHERE"
    strSQL = strSQL & " MORTGAGEPRODUCTCODE = " & objSqlAssist.FormatString(strMortgageProductCode)
    strSQL = strSQL & " AND STARTDATE = " & objSqlAssist.FormatDate(strStartDate)
    strSQL = strSQL & " AND INCENTIVE.INCENTIVEGUID = INCLUSIVEINCENTIVE.INCENTIVEGUID"

    cmd.ActiveConnection = objIADOAssist.GetConnStr
    cmd.CommandText = strSQL
    
    Set rstThisRecSet = cmd.Execute(1, , adCmdText)
    
    If rstThisRecSet.EOF Then
        ' raise application error to be interpreted by calling object
        m_objErrAssist.RaiseError TypeName(Me), strFunctionName, omiga4RecordNotFound
    End If
    
    Dim objXmlElem As IXMLDOMElement
    Dim objXmlIndicatorElem As IXMLDOMElement
    Dim objXmlListNode As IXMLDOMNode
    Dim objXmlOut As FreeThreadedDOMDocument40
    Dim objXmlNode As IXMLDOMNode
    
    Dim objTempElement As IXMLDOMElement
    
    Dim dblPercentage As Double
    Dim dblAmount As Double
    Dim dblLoanAmount As Double
    Dim dblPercentageMaximum As Double
    Dim objConvertAssist As ConvertAssist
    
    Set objXmlOut = New FreeThreadedDOMDocument40
    objXmlOut.validateOnParse = False
    objXmlOut.setProperty "NewParser", True
    Set objXmlElem = objXmlOut.createElement("INCLUSIVEINCENTIVELIST")
    Set objXmlListNode = objXmlOut.appendChild(objXmlElem)
        
    'BG 17/10/00 SYS1600
    Dim xmlLoadIncentiveData As FreeThreadedDOMDocument40
    
    Dim objIomMpClassDef As IomMPClassDef
    Set objIomMpClassDef = New omMPClassDef
        
    Set xmlLoadIncentiveData = objIomMpClassDef.LoadIncentiveData
    'BG 17/10/00 End SYS1600
    
    Set objConvertAssist = New ConvertAssist
    
    ' AD 08/10/01. Removed the MoveFirst as it's a ForwardOnly cursor. SYS2718
       ' loop through the record set
    While Not rstThisRecSet.EOF
        
        Set objXmlNode = m_objIDOAssist.GetXMLFromRecordSet(rstThisRecSet, xmlLoadIncentiveData)
        
        dblPercentage = objConvertAssist.CSafeDbl(m_objXmlAssist.GetTagValue(objXmlNode, "PERCENTAGE"))
        dblLoanAmount = objConvertAssist.CSafeDbl(strLoanAmount)
        
        'DPF 21/10/02 - CPWP1 add new field
        Set objXmlIndicatorElem = objXmlOut.createElement("INCENTIVEFIXEDINDICATOR")
        
        If dblPercentage > 0 Then
            Set objTempElement = objXmlNode.ownerDocument.getElementsByTagName("AMOUNT").Item(0)
            dblAmount = ((dblPercentage / 100) * dblLoanAmount)
            dblPercentageMaximum = objConvertAssist.CSafeDbl(m_objXmlAssist.GetTagValue(objXmlNode, "PERCENTAGEMAXIMUM"))
            
            objTempElement.Text = CStr(dblAmount)
            
            If dblPercentageMaximum <> 0 And dblAmount > dblPercentageMaximum Then
                objTempElement.Text = CStr(dblPercentageMaximum)
            End If
            
            'DPF 21/10/02 - CPWP1 add new field
            objXmlIndicatorElem.Text = "0"
            objXmlNode.appendChild objXmlIndicatorElem
        Else
            'DPF 21/10/02 - CPWP1 add new field
            objXmlIndicatorElem.Text = "1"
            objXmlNode.appendChild objXmlIndicatorElem
        End If
                
        objXmlListNode.appendChild objXmlNode
        
        rstThisRecSet.MoveNext
    Wend
    rstThisRecSet.Close
    GetInclusiveIncentivesForProduct = objXmlOut.xml
    
    Set objSqlAssist = Nothing
    Set rstThisRecSet = Nothing
    Set cmd = Nothing
    Set objIADOAssist = Nothing
    Set objXmlIn = Nothing

    If Not m_objContext Is Nothing Then
        m_objContext.SetComplete
    End If
    
    Exit Function

GetInclusiveIncentivesForProductVbErr:
        
    ' Close the recordset if still open
    If Not rstThisRecSet Is Nothing Then
        If rstThisRecSet.State = adStateOpen Then
            rstThisRecSet.Close
        End If
    End If

    Set rstThisRecSet = Nothing
    Set cmd = Nothing
    
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

Public Function GetMIGRate(ByVal vxmlRequest As IXMLDOMElement) As IXMLDOMNode
    
    On Error GoTo GetMIGRateVbErr

    Const strFunctionName As String = "GetMIGRate"
    
    Dim objSqlAssist As SQLAssist
    Set objSqlAssist = New SQLAssist
    
    Dim objIADOAssist As IADOAssist
        
    Dim strInLTV As String
    Dim strInMaximumLoanAmount As String
    Dim strInRateSet As String
    Dim strApplicationDate As String
    
    strInRateSet = m_objXmlAssist.GetTagValue(vxmlRequest, "RATESET")
    strInMaximumLoanAmount = m_objXmlAssist.GetTagValue(vxmlRequest, "TOTALLOANAMOUNT")
    strInLTV = m_objXmlAssist.GetTagValue(vxmlRequest, "LTV")
    strApplicationDate = m_objXmlAssist.GetTagValue(vxmlRequest, "APPLICATIONDATE")
    
    'BMIDS821 GHun Check that ApplicationDate is a date
    If Len(strInRateSet) = 0 Or _
       Len(strInMaximumLoanAmount) = 0 Or _
       Len(strInLTV) = 0 Or _
       Not (IsDate(strApplicationDate)) Then
       Dim strDetails As String
       strDetails = "MPMIGRATESET = " & strInRateSet
       strDetails = strDetails & "; AMOUNTREQUESTED = " & strInMaximumLoanAmount
       strDetails = strDetails & "; LTV = " & strInLTV
       'BMIDS821 GHun
       strDetails = strDetails & "; APPLICATIONDATE = " & strApplicationDate
       
       m_objErrAssist.RaiseError TypeName(Me), _
                               strFunctionName, _
                               omiga4InvalidParameter, _
                               strDetails & "  No applicable MIGRateSet for the product"
    End If
                
    If m_objContext Is Nothing Then
        Set objIADOAssist = New ADOAssist
    Else
        Set objIADOAssist = m_objContext.CreateInstance(gstrBASE_COMPONENT & ".ADOAssist")
    End If
    
    'Build SQL string
    
    'Dim strSQL As String
    'strSQL = "SELECT * FROM MIGRATE MIGRATE1,MIGRATESET"
    'strSQL = strSQL & " WHERE MIGRATESET.RATESET = " & strInRateSet
    'strSQL = strSQL & " AND MIGRATE1.MIGRATESTARTDATE = (SELECT MAX(MIGRATESTARTDATE) FROM MIGRATE"
    'strSQL = strSQL & " WHERE MIGRATESTARTDATE <= " & objSQLAssist.GetSystemDate
    'strSQL = strSQL & " AND MIGRATE.RATESET = MIGRATESET.RATESET " & ")"
    'strSQL = strSQL & " AND MIGRATE1.MAXIMUMLTV = (SELECT MIN(MAXIMUMLTV) FROM MIGRATE"
    'strSQL = strSQL & " WHERE MAXIMUMLTV >= " & strInLTV & " AND MIGRATE.RATESET = MIGRATESET.RATESET"
    'strSQL = strSQL & " AND MIGRATESTARTDATE = MIGRATE1.MIGRATESTARTDATE )"
    'strSQL = strSQL & " AND MIGRATE1.MAXIMUMLOANAMOUNT = (SELECT MIN(MAXIMUMLOANAMOUNT) FROM MIGRATE"
    'strSQL = strSQL & " WHERE MAXIMUMLOANAMOUNT >= " & strInMaximumLoanAmount
    'strSQL = strSQL & " AND MIGRATE.RATESET = MIGRATESET.RATESET"
    'strSQL = strSQL & " AND MIGRATESTARTDATE = MIGRATE1.MIGRATESTARTDATE)"
    'strSQL = strSQL & " AND MIGRATE1.RATESET = MIGRATESET.RATESET "
    
    'MO - 22/10/2002 - BMIDS00666 - Start
    'Dim strSQL As String
    'strSQL = "SELECT * FROM  MPMIGRATESET , MPMIGRATEBAND "
    'strSQL = strSQL + " WHERE MPMIGRATESET.MPMIGRATESET  = MPMIGRATEBAND.MPMIGRATESET "
    'strSQL = strSQL + " AND MPMIGRATESET.MPMIGRATESET = " & strInRateSet
    'strSQL = strSQL + " AND MPMIGRATEBAND.EFFECTIVEDATE = (SELECT MAX(EFFECTIVEDATE) FROM MPMIGRATEBAND "
    'strSQL = strSQL + " WHERE EFFECTIVEDATE <= " & objSQLAssist.GetSystemDate
    'strSQL = strSQL + " AND MPMIGRATESET.MPMIGRATESET = MPMIGRATEBAND.MPMIGRATESET ) "
    'strSQL = strSQL + " AND MPMIGRATEBAND.ENDDATE = ( SELECT MIN(ENDDATE) FROM MPMIGRATEBAND "
    'strSQL = strSQL + " WHERE ENDDATE <= " & objSQLAssist.GetSystemDate
    'strSQL = strSQL + " AND MPMIGRATESET.MPMIGRATESET = MPMIGRATEBAND.MPMIGRATESET) "
    'strSQL = strSQL + " AND MPMIGRATEBAND.LOWERLTV  = (SELECT MAX(LOWERLTV) FROM MPMIGRATEBAND "
    'strSQL = strSQL + " WHERE LOWERLTV <= " & strInLTV
    'strSQL = strSQL + " AND MPMIGRATESET.MPMIGRATESET = MPMIGRATEBAND.MPMIGRATESET) "
    'strSQL = strSQL + " AND MPMIGRATEBAND.HIGHERLTV  = (SELECT MIN(HIGHERLTV) FROM MPMIGRATEBAND "
    'strSQL = strSQL + " WHERE HIGHERLTV >= " & strInLTV
    'strSQL = strSQL + " AND MPMIGRATESET.MPMIGRATESET = MPMIGRATEBAND.MPMIGRATESET) "
    'strSQL = strSQL + " AND MPMIGRATEBAND.LOWERLOANAMOUNT = ( SELECT MAX(LOWERLOANAMOUNT) FROM MPMIGRATEBAND "
    'strSQL = strSQL + " WHERE LOWERLOANAMOUNT <= " & strInMaximumLoanAmount
    'strSQL = strSQL + " AND MPMIGRATESET.MPMIGRATESET = MPMIGRATEBAND.MPMIGRATESET) "
    'strSQL = strSQL + " AND MPMIGRATEBAND.HIGHERLOANAMOUNT = ( SELECT MIN(HIGHERLOANAMOUNT) FROM MPMIGRATEBAND "
    'strSQL = strSQL + " WHERE HIGHERLOANAMOUNT >= " & strInMaximumLoanAmount
    'strSQL = strSQL + " AND MPMIGRATESET.MPMIGRATESET = MPMIGRATEBAND.MPMIGRATESET) "
    
    Dim strSQL As String
    
    '*=[MC] BMIDS763 - GET most recent MIG Rate which is valid with the given Product Start date.
    strSQL = "SELECT TOP 1 * " & _
                "FROM  MPMIGRATESET " & _
                "INNER JOIN MPMIGRATEBAND ON MPMIGRATESET.MPMIGRATESET = MPMIGRATEBAND.MPMIGRATESET " & _
                "WHERE MPMIGRATESET.MPMIGRATESET = " & strInRateSet & " " & _
                  " AND MPMIGRATEBAND.EFFECTIVEDATE<= " & objSqlAssist.FormatDate(strApplicationDate, dtfDate) & _
                  " AND MPMIGRATEBAND.ENDDATE >= " & objSqlAssist.FormatDate(strApplicationDate, dtfDate) & _
                  " AND MPMIGRATEBAND.LOWERLTV <= " & strInLTV & " " & _
                  " AND MPMIGRATEBAND.HIGHERLTV >= " & strInLTV & " " & _
                  " AND MPMIGRATEBAND.LOWERLOANAMOUNT <= " & strInMaximumLoanAmount & " " & _
                  " AND MPMIGRATEBAND.HIGHERLOANAMOUNT >= " & strInMaximumLoanAmount & "  " & " ORDER BY MPMIGRATEBAND.EFFECTIVEDATE DESC "
    
    Dim cmd As ADODB.Command
    Set cmd = New ADODB.Command
    
    cmd.ActiveConnection = objIADOAssist.GetConnStr
    cmd.CommandText = strSQL
    
    Dim rstThisRecSet As ADODB.Recordset
    Set rstThisRecSet = New ADODB.Recordset
    
    Set rstThisRecSet = cmd.Execute(1, , adCmdText)
            
    Dim xmlLoadMPMigRateData As FreeThreadedDOMDocument40
    Dim IOmOrgClassDef As IOmOrgClassDef
    Set IOmOrgClassDef = New omOrgClassDef
    Set xmlLoadMPMigRateData = IOmOrgClassDef.LoadMPMigRateData()
    
    '*=[MC]BMIDS763
    If rstThisRecSet.BOF And rstThisRecSet.EOF Then
        strSQL = "SELECT TOP 1 * " & _
                "FROM  MPMIGRATESET " & _
                "INNER JOIN MPMIGRATEBAND ON MPMIGRATESET.MPMIGRATESET = MPMIGRATEBAND.MPMIGRATESET " & _
                "WHERE MPMIGRATESET.MPMIGRATESET = " & strInRateSet & " " & _
                  "  AND MPMIGRATEBAND.EFFECTIVEDATE<= " & objSqlAssist.GetSystemDate & _
                  "  AND MPMIGRATEBAND.ENDDATE >= " & objSqlAssist.GetSystemDate & _
                  "AND MPMIGRATEBAND.LOWERLTV <= " & strInLTV & " " & _
                  "AND MPMIGRATEBAND.HIGHERLTV >= " & strInLTV & " " & _
                  "AND MPMIGRATEBAND.LOWERLOANAMOUNT <= " & strInMaximumLoanAmount & " " & _
                  "AND MPMIGRATEBAND.HIGHERLOANAMOUNT >= " & strInMaximumLoanAmount & "  ORDER BY MPMIGRATEBAND.EFFECTIVEDATE DESC "
        cmd.CommandText = strSQL
        Set rstThisRecSet = cmd.Execute(1, , adCmdText)
    End If
    
    If Not rstThisRecSet.EOF Then
        ' convert record set to xml
        Set GetMIGRate = m_objIDOAssist.GetXMLFromRecordSet(rstThisRecSet, xmlLoadMPMigRateData)
    Else
        m_objErrAssist.RaiseError TypeName(Me), strFunctionName, omiga4RecordNotFound, "No applicable MIGRate/MIGRateBand is found. Please contact Help Desk"
    End If
    
    rstThisRecSet.Close

    Set cmd = Nothing
    Set rstThisRecSet = Nothing
    Set objIADOAssist = Nothing
    
    If Not m_objContext Is Nothing Then
        m_objContext.SetComplete
    End If

    Exit Function

GetMIGRateVbErr:
    
    ' Close the recordset if still open
    If Not rstThisRecSet Is Nothing Then
        If rstThisRecSet.State = adStateOpen Then
            rstThisRecSet.Close
        End If
    End If

    Set cmd = Nothing
    Set rstThisRecSet = Nothing
    Set objIADOAssist = Nothing
    
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

Public Function AddDerivedData(ByVal vstrXMLData As String) As String
' header ----------------------------------------------------------------------------------
' description:
'   XML elements must be created for any derived values as specified.
'   Add any derived values to XML. E.g. data type 'double' fields will
'   need to be formatted as strings to required precision & rounding.
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
On Error GoTo AddDerivedDataVBErr

    Const strFunctionName As String = "AddDerivedData"

' TODO ************************************************************************
'   Remove omiga4NotImplemented error when this function has been implemented
' *****************************************************************************
    m_objErrAssist.RaiseError TypeName(Me), strFunctionName, omiga4NotImplemented
        
    AddDerivedData = vstrXMLData
       
    Exit Function
    
AddDerivedDataVBErr:

    AddDerivedData = ""

    Err.Raise Err.Number, Err.Source, Err.Description, Err.HelpFile, Err.HelpContex
    
End Function

Private Function IMortgageProductDO_FindProductList(ByVal vxmlTableElement As IXMLDOMElement) As IXMLDOMNode
' header ----------------------------------------------------------------------------------
' description:
'   Get the data for a single instance of the persistant data associated with
'   this data object
' pass:
'   vxmlxmlTableElement  xml element containing the request
' return:                xml node containing retrieved data
'------------------------------------------------------------------------------------------
On Error GoTo FindProductListVbErr

    Const strFunctionName As String = "IMortgageProductDO_FindProductList"
    
    Dim xmlOut As FreeThreadedDOMDocument40
    Dim objSqlAssist As SQLAssist
    Dim xmlListElem As IXMLDOMElement
    Dim xmlClassDefDoc As FreeThreadedDOMDocument40
    Dim adoConnection As ADODB.Connection
    Dim adoCommand As ADODB.Command
    Dim rstThisRecSet As ADODB.Recordset
    Dim adoParameter As ADODB.Parameter
    Dim objIomMpClassDef As IomMPClassDef
    Dim objGlobalParamDO As GlobalParameterDO

    Dim objADOAssist As ADOAssist
                
    If vxmlTableElement Is Nothing Then
        m_objErrAssist.ThrowError TypeName(Me), _
                                strFunctionName, _
                                oeInvalidParameter, _
                                "Request node is nothing"
    End If
           
    Dim strMaxRecords As String
    Dim strCountryOfOrigin As String
    Dim strLanguage As String
    Dim strStartRecord As String
    Dim strRecordCount As String
    Dim strTypeOfApplication As String
    Dim strAmountRequested As String
    Dim strTermInMonths As String
    Dim strTermInYears As String
    Dim strLTV As String
    Dim strRepaymentMethod As String
    Dim strInterestOnlyElement As String
    Dim strCapitalAndInterestElement As String
    Dim strPurposeOfLoan As String
    Dim strLenderId As String
    Dim strDiscountedProducts As String
    Dim strDiscountedPeriod As String
    Dim strFixedRateProducts As String
    Dim strFixedRatePeriod As String
    Dim strStandardVariableProducts As String
    Dim strCappedFlooredProducts As String
    Dim strCappedFlooredPeriod As String
    Dim strCashbackProducts As String
    'Dim strFlexibleProducts As String  'EP2_9 GHun
    Dim strImpairedCreditRating As String
    Dim strProductsByGroup As String
    Dim strProductGroup As String
    Dim strDistributionChannelId As String
    Dim strAllProductsWithChecks As String
    Dim strSearchContext As String
    Dim strMemberOfStaff As String
    Dim strExistingCustomer As String
    'Dim strSelfEmployed As String      '--BMIDS00246
    Dim strTypeOfBuyer As String
    Dim strApplicationNumber As String
    'CMWP3 - DPF - Jul 02 new variable as we are going to allow searching by portable loan indicator
    Dim strManualPortedLoanInd As String
        
    'AW BMIDS00077 - MSMS Integration
    Dim strMortgageProductCode As String    'AW 25/06/02 SYS4202
    'AW End
    
    Dim strEmploymentStatusList As String 'SA BMIDS00246
    
    Dim strOrderBy As String
    Dim strMaxAge As String 'MAR1702
    Dim strMinAge As String 'MAR1702
    Dim strApp1DOB As String 'MAR1702
    Dim strApp2DOB As String 'MAR1702
    Dim intAge As Integer   'MAR1878 GHun
    
    'EP2_9 GHun
    Dim strNatureOfLoan As String
    Dim strIncomeStatus As String
    Dim strCreditScheme As String
    Dim strGuarantorInd As String
    Dim strFlexNonFlex  As String
    Dim strFlexibleInd  As String
    'EP2_9 End
    
    'EP2_76 SR : 13/11/2006
    Dim strPrincipalFirmId As String
    Dim strARFirmId As String
    Dim strMortgageClubId As String
    Dim strPackagerId As String
    Dim strPackagerAssociationId As String
    'EP2_76 SR  - End
    
    strCountryOfOrigin = m_objXmlAssist.GetTagValue(vxmlTableElement, "COUNTRYOFORIGIN")
    strLanguage = m_objXmlAssist.GetTagValue(vxmlTableElement, "LANGUAGE")
    strTypeOfApplication = m_objXmlAssist.GetTagValue(vxmlTableElement, "TYPEOFAPPLICATION")
    strAmountRequested = m_objXmlAssist.GetTagValue(vxmlTableElement, "AMOUNTREQUESTED")
    strTermInYears = m_objXmlAssist.GetTagValue(vxmlTableElement, "TERMINYEARS")
    strTermInMonths = m_objXmlAssist.GetTagValue(vxmlTableElement, "TERMINMONTHS")
    strLTV = m_objXmlAssist.GetTagValue(vxmlTableElement, "LTV")
    strRepaymentMethod = m_objXmlAssist.GetTagValue(vxmlTableElement, "REPAYMENTMETHOD")
    strInterestOnlyElement = m_objXmlAssist.GetTagValue(vxmlTableElement, "INTERESTONLYELEMENT")
    strCapitalAndInterestElement = m_objXmlAssist.GetTagValue(vxmlTableElement, "CAPITALANDINTERESTELEMENT")
    strPurposeOfLoan = m_objXmlAssist.GetTagValue(vxmlTableElement, "PURPOSEOFLOAN")
    strLenderId = m_objXmlAssist.GetTagValue(vxmlTableElement, "LENDERID")
    strDiscountedProducts = m_objXmlAssist.GetTagValue(vxmlTableElement, "DISCOUNTEDPRODUCTS")
    strDiscountedPeriod = m_objXmlAssist.GetTagValue(vxmlTableElement, "DISCOUNTEDPERIOD")
    strFixedRateProducts = m_objXmlAssist.GetTagValue(vxmlTableElement, "FIXEDRATEPRODUCTS")
    strFixedRatePeriod = m_objXmlAssist.GetTagValue(vxmlTableElement, "FIXEDRATEPERIOD")
    strStandardVariableProducts = m_objXmlAssist.GetTagValue(vxmlTableElement, "STANDARDVARIABLEPRODUCTS")
    strCappedFlooredProducts = m_objXmlAssist.GetTagValue(vxmlTableElement, "CAPPEDFLOOREDPRODUCTS")
    strCappedFlooredPeriod = m_objXmlAssist.GetTagValue(vxmlTableElement, "CAPPEDFLOOREDPERIOD")
    strCashbackProducts = m_objXmlAssist.GetTagValue(vxmlTableElement, "CASHBACKPRODUCTS")
    'strFlexibleProducts = m_objXmlAssist.GetTagValue(vxmlTableElement, "FLEXIBLEPRODUCTS")
    strImpairedCreditRating = m_objXmlAssist.GetTagValue(vxmlTableElement, "IMPAIREDCREDITRATING")
    strProductsByGroup = m_objXmlAssist.GetTagValue(vxmlTableElement, "PRODUCTSBYGROUP")
    strProductGroup = m_objXmlAssist.GetTagValue(vxmlTableElement, "PRODUCTGROUP")
    strDistributionChannelId = m_objXmlAssist.GetTagValue(vxmlTableElement, "DISTRIBUTIONCHANNELID")
    strAllProductsWithChecks = m_objXmlAssist.GetTagValue(vxmlTableElement, "ALLPRODUCTSWITHCHECKS")
    strSearchContext = m_objXmlAssist.GetTagValue(vxmlTableElement, "SEARCHCONTEXT")
    strMemberOfStaff = m_objXmlAssist.GetTagValue(vxmlTableElement, "MEMBEROFSTAFF")
    strExistingCustomer = m_objXmlAssist.GetTagValue(vxmlTableElement, "EXISTINGCUSTOMER")
    strTypeOfBuyer = m_objXmlAssist.GetTagValue(vxmlTableElement, "TYPEOFBUYER")
    
    'CMWP3 - DPF - Jul 02 - new variable added for searching on
    strManualPortedLoanInd = m_objXmlAssist.GetTagValue(vxmlTableElement, "MANUALPORTEDLOANIND")
    strApplicationNumber = m_objXmlAssist.GetTagValue(vxmlTableElement, "APPLICATIONNUMBER")
    strStartRecord = m_objXmlAssist.GetTagValue(vxmlTableElement, "STARTRECORD")
    strRecordCount = m_objXmlAssist.GetTagValue(vxmlTableElement, "RECORDCOUNT")
    
    'AW BMIDS00077 - MSMS Integration
    'AW 25/06/02 SYS4202
    strMortgageProductCode = m_objXmlAssist.GetTagValue(vxmlTableElement, "MORTGAGEPRODUCTCODE")
    'AW End
    
    'BMIDS00246 SA {
    strEmploymentStatusList = m_objXmlAssist.GetTagValue(vxmlTableElement, "EMPLOYMENTSTATUSLIST")
    'BMIDS00246 }
        
    'MAR1702
    strApp1DOB = m_objXmlAssist.GetTagValue(vxmlTableElement, "APPLICANT1DOB")
    strApp2DOB = m_objXmlAssist.GetTagValue(vxmlTableElement, "APPLICANT2DOB")

    'MAR1702
    strMaxAge = vbNullString
    strMinAge = vbNullString
    
    If Len(strApp1DOB) > 0 Then
        strMaxAge = CStr(GetAge(strApp1DOB))
        strMinAge = strMaxAge   'MAR1878 GHun
    End If
        
    'MAR1702
    If Len(strApp2DOB) > 0 Then
        'MAR1878 GHun Only calculate age once, and then reuse the calculated value
        intAge = GetAge(strApp2DOB)
    
        If Len(strMaxAge) = 0 Then
            strMaxAge = CStr(intAge)
        Else
            If (intAge > CInt(strMaxAge)) Then
                strMaxAge = CStr(intAge)
            End If
        End If
        If Len(strMinAge) = 0 Then
            strMinAge = CStr(intAge)
        Else
            If intAge < CInt(strMinAge) Or (Len(strMinAge) = 0) Then
                strMinAge = CStr(intAge)
            End If
        End If
        'MAR1878 End
    End If
    
    If Len(strMaxAge) = 0 Then
        strMaxAge = "0"
    End If
    If Len(strMinAge) = 0 Then
        strMinAge = "0"
    End If
    
    'EP2_9 GHun
    strNatureOfLoan = m_objXmlAssist.GetTagValue(vxmlTableElement, "NATUREOFLOAN")
    strIncomeStatus = m_objXmlAssist.GetTagValue(vxmlTableElement, "APPLICATIONINCOMESTATUS")
    strCreditScheme = m_objXmlAssist.GetTagValue(vxmlTableElement, "CREDITSCHEME")
    'CreditScheme takes precedence over ProductGroup
    If Len(strCreditScheme) > 0 Then
        strProductsByGroup = "1"
        strProductGroup = strCreditScheme
    End If
    strGuarantorInd = m_objXmlAssist.GetTagValue(vxmlTableElement, "GUARANTORIND")
    strFlexNonFlex = m_objXmlAssist.GetTagValue(vxmlTableElement, "FLEXIBLENONFLEXIBLEPRODUCTS")
    strFlexibleInd = m_objXmlAssist.GetTagValue(vxmlTableElement, "FLEXIBLEIND")
    'EP2_9 End
    
    'SR 13/11/2006 : EP2_76
    strPrincipalFirmId = m_objXmlAssist.GetTagValue(vxmlTableElement, "PRINCIPALFIRMID")
    strARFirmId = m_objXmlAssist.GetTagValue(vxmlTableElement, "ARFIRMID")
    strMortgageClubId = m_objXmlAssist.GetTagValue(vxmlTableElement, "MORTGAGECLUBID")
    strPackagerId = m_objXmlAssist.GetTagValue(vxmlTableElement, "PACKAGERID")
    strPackagerAssociationId = m_objXmlAssist.GetTagValue(vxmlTableElement, "PACKAGERASSOCIATIONID")
    'SR 13/11/2006 : EP2_76 - End
    
    ' If entered, use the number of records asked for as the maximum records
    If Len(strRecordCount) > 0 And IsNumeric(strRecordCount) Then
        strMaxRecords = strRecordCount
    Else
        ' Get the default number of records
        If m_objContext Is Nothing Then
            Set objGlobalParamDO = New GlobalParameterDO
        Else
            Set objGlobalParamDO = m_objContext.CreateInstance(gstrBASE_COMPONENT & ".GlobalParameterDO")
        End If
            
        On Error Resume Next
        strMaxRecords = objGlobalParamDO.GetCurrentParameterByType("MortgageProductMaxRecords", "AMOUNT")
        
        If m_objErrAssist.GetOmigaErrorNumber(Err.Number) = oeRecordNotFound Then
            m_objErrAssist.ThrowError TypeName(Me), strFunctionName, 213
        End If
        
        On Error GoTo FindProductListVbErr
    
    End If
        
    strOrderBy = m_objXmlAssist.GetTagValue(vxmlTableElement, "ORDERBY")
    
    If m_objContext Is Nothing Then
        Set objADOAssist = New ADOAssist
    Else
        Set objADOAssist = m_objContext.CreateInstance(gstrBASE_COMPONENT & ".ADOAssist")
    End If
    
    Set adoConnection = New ADODB.Connection
    With adoConnection
        .ConnectionString = objADOAssist.GetConnStr
        .CursorLocation = adUseClient
        .Open
    End With
    
    Set rstThisRecSet = New ADODB.Recordset
    rstThisRecSet.CursorLocation = adUseClient
    
    Set objSqlAssist = New SQLAssist
        
    'MAR1878 GHun Call usp_MPSFindProducts as a stored proc, not text
    '' AS 12/06/01 CC012 - Start
    'Dim eDbEngineType As DBENGINETYPE
    'eDbEngineType = objADOAssist.GetDBEngine
    'If eDbEngineType = SQLServer Then
    '    'BMIDS00246 Added extra parameter and then removed original IsSelfEmployed parameter
    '    strSQL = "{CALL usp_MPSFindProducts("
    '    strSQL = strSQL & "?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,"
    '    strSQL = strSQL & "?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?)}"
    'ElseIf eDbEngineType = Oracle Then
    '    strSQL = "{CALL MORTGAGEPRODUCTSELECTION.FINDPRODUCTS("
    '    strSQL = strSQL & "?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,"
    '    strSQL = strSQL & "?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,"
    '    strSQL = strSQL & "{RESULTSET " & strMaxRecords & ","
    '    'strSQL = strSQL & "MORTGAGEPRODUCTCODE,STARTDATE,PRODUCTNAME,PRODUCTTEXTDETAILS,"
    '    'AW BMIDS00077 - MSMS Integration
    '    'AW 25/06/02 SYS4204 Added OrganisationID to output
    '    strSQL = strSQL & "MORTGAGEPRODUCTCODE,STARTDATE,PRODUCTNAME,PRODUCTTEXTDETAILS,ORGANISATIONID,"
    '    'AW End
    '    strSQL = strSQL & "LENDERSNAME,FIRSTMONTHSREPAYMENT,FIRSTMONTHSREPAYMENTBYDAILY,FIRSTMONTHLYINTERESTRATE,"
    '    strSQL = strSQL & "DISCOUNTEDPERIOD,TYPE,FIRSTMONTHLYSVRREPAYMENT,FIRSTMONTHLYSVRREPAYBYDAILY,FEE,PENALTIES,"
    '    strSQL = strSQL & "FLEXIBLEMORTGAGEPRODUCT,NONPANELLENDEROPTION,"
    '    strSQL = strSQL & "INTERESTRATEENDDATE,INTERESTRATEPERIOD,TYPICALAPR})}"
    'End If
    '' AS 12/06/01 CC012 - End

    Set adoCommand = New ADODB.Command
    With adoCommand
        Set .ActiveConnection = adoConnection
        .CommandType = adCmdStoredProc
        .CommandText = "dbo.usp_MPSFindProducts"
    End With
    'MAR1878 End

    Set adoParameter = adoCommand.CreateParameter("@p_iCountryOfOrigin", adInteger, adParamInput)
    adoParameter.Attributes = adParamNullable
    If Len(strCountryOfOrigin) = 0 Then
        adoParameter.Value = Null
    ElseIf IsNumeric(strCountryOfOrigin) Then
        adoParameter.Value = strCountryOfOrigin
    Else
        m_objErrAssist.ThrowError TypeName(Me), _
                                strFunctionName, _
                                oeInvalidParameter, _
                                "COUNTRYOFORIGIN must be numeric or empty"
    End If
    adoCommand.Parameters.Append adoParameter
    
    Set adoParameter = adoCommand.CreateParameter("@p_iLanguage", adInteger, adParamInput)
    adoParameter.Attributes = adParamNullable
    If Len(strLanguage) = 0 Then
        adoParameter.Value = Null
    ElseIf IsNumeric(strLanguage) Then
        adoParameter.Value = strLanguage
    Else
        m_objErrAssist.ThrowError TypeName(Me), _
                                strFunctionName, _
                                oeInvalidParameter, _
                                "LANGUAGE must be numeric or empty"
    End If
    adoCommand.Parameters.Append adoParameter

    Set adoParameter = adoCommand.CreateParameter("@p_iStartRecord", adInteger, adParamInput)
    adoParameter.Attributes = adParamNullable
    If Len(strStartRecord) = 0 Then
        adoParameter.Value = Null
    ElseIf IsNumeric(strStartRecord) Then
        adoParameter.Value = strStartRecord
    Else
        m_objErrAssist.ThrowError TypeName(Me), _
                                strFunctionName, _
                                oeInvalidParameter, _
                                "STARTRECORD must be numeric or empty"
    End If
    adoCommand.Parameters.Append adoParameter

    adoCommand.Parameters.Append adoCommand.CreateParameter("@p_iRecordCount", adInteger, adParamInput, , CLng(strMaxRecords))

    Set adoParameter = adoCommand.CreateParameter("@p_iTypeOfApplication", adInteger, adParamInput)
    adoParameter.Attributes = adParamNullable
    If Len(strTypeOfApplication) = 0 Then
        adoParameter.Value = Null
    ElseIf IsNumeric(strTypeOfApplication) Then
        adoParameter.Value = strTypeOfApplication
    Else
        m_objErrAssist.ThrowError TypeName(Me), _
                                strFunctionName, _
                                oeInvalidParameter, _
                                "TYPEOFAPPLICATION must be numeric or empty"
    End If
    adoCommand.Parameters.Append adoParameter

    Set adoParameter = adoCommand.CreateParameter("@p_iAmountRequested", adInteger, adParamInput)
    adoParameter.Attributes = adParamNullable
    If Len(strAmountRequested) = 0 Then
        adoParameter.Value = Null
    ElseIf IsNumeric(strAmountRequested) Then
        adoParameter.Value = strAmountRequested
    Else
        m_objErrAssist.ThrowError TypeName(Me), _
                                strFunctionName, _
                                oeInvalidParameter, _
                                "AMOUNTREQUESTED must be numeric or empty"
    End If
    adoCommand.Parameters.Append adoParameter

    Set adoParameter = adoCommand.CreateParameter("@p_iTermInYears", adInteger, adParamInput)
    adoParameter.Attributes = adParamNullable
    If Len(strTermInYears) = 0 Then
        adoParameter.Value = Null
    ElseIf IsNumeric(strTermInYears) Then
        adoParameter.Value = strTermInYears
    Else
        m_objErrAssist.ThrowError TypeName(Me), _
                                strFunctionName, _
                                oeInvalidParameter, _
                                "TERMINYEARS must be numeric or empty"
    End If
    adoCommand.Parameters.Append adoParameter

    Set adoParameter = adoCommand.CreateParameter("@p_iTermInMonths", adInteger, adParamInput)
    adoParameter.Attributes = adParamNullable
    If Len(strTermInMonths) = 0 Then
        adoParameter.Value = Null
    ElseIf IsNumeric(strTermInMonths) Then
        adoParameter.Value = strTermInMonths
    Else
        m_objErrAssist.ThrowError TypeName(Me), _
                                strFunctionName, _
                                oeInvalidParameter, _
                                "TERMINMONTHS must be numeric or empty"
    End If
    adoCommand.Parameters.Append adoParameter

    Set adoParameter = adoCommand.CreateParameter("@p_iLTV", adNumeric, adParamInput)
    adoParameter.NumericScale = 2
    'AW BMIDS00077 - MSMS Integration
'   adoParameter.Precision = 9
    adoParameter.Precision = 7
    'AW End
    adoParameter.Attributes = adParamNullable
    If Len(strLTV) = 0 Then
        adoParameter.Value = Null
    ElseIf IsNumeric(strLTV) Then
        adoParameter.Value = strLTV
    Else
        m_objErrAssist.ThrowError TypeName(Me), _
                                strFunctionName, _
                                oeInvalidParameter, _
                                "LTV must be numeric or empty"
    End If
    adoCommand.Parameters.Append adoParameter
    
    'EP2_9 GHun
    'Set adoParameter = adoCommand.CreateParameter("@p_iRepaymentMethod", adInteger, adParamInput)
    'adoParameter.Attributes = adParamNullable
    'If Len(strRepaymentMethod) = 0 Then
    '    adoParameter.Value = Null
    'ElseIf IsNumeric(strRepaymentMethod) Then
    '    adoParameter.Value = strRepaymentMethod
    'Else
    '    m_objErrAssist.ThrowError TypeName(Me), _
    '                            strFunctionName, _
    '                            oeInvalidParameter, _
    '                            "REPAYMENTMETHOD must be numeric or empty"
    'End If
    'adoCommand.Parameters.Append adoParameter
    'EP2_9 End

    Set adoParameter = adoCommand.CreateParameter("@p_iInterestOnlyElement", adInteger, adParamInput)
    adoParameter.Attributes = adParamNullable
    If Len(strInterestOnlyElement) = 0 Then
        adoParameter.Value = Null
    ElseIf IsNumeric(strInterestOnlyElement) Then
        adoParameter.Value = strInterestOnlyElement
    Else
        m_objErrAssist.ThrowError TypeName(Me), _
                                strFunctionName, _
                                oeInvalidParameter, _
                                "INTERESTONLYELEMENT must be numeric or empty"
    End If
    adoCommand.Parameters.Append adoParameter
    
    Set adoParameter = adoCommand.CreateParameter("@p_iCapitalAndInterestElement", adInteger, adParamInput)
    adoParameter.Attributes = adParamNullable
    If Len(strCapitalAndInterestElement) = 0 Then
        adoParameter.Value = Null
    ElseIf IsNumeric(strCapitalAndInterestElement) Then
        adoParameter.Value = strCapitalAndInterestElement
    Else
        m_objErrAssist.ThrowError TypeName(Me), _
                                strFunctionName, _
                                oeInvalidParameter, _
                                "CAPITALANDINTERESTELEMENT must be numeric or empty"
    End If
    adoCommand.Parameters.Append adoParameter
    
    Set adoParameter = adoCommand.CreateParameter("@p_iPurposeOfLoan", adInteger, adParamInput)
    adoParameter.Attributes = adParamNullable
    If Len(strPurposeOfLoan) = 0 Then
        adoParameter.Value = Null
    ElseIf IsNumeric(strPurposeOfLoan) Then
        adoParameter.Value = strPurposeOfLoan
    Else
        m_objErrAssist.ThrowError TypeName(Me), _
                                strFunctionName, _
                                oeInvalidParameter, _
                                "PURPOSEOFLOAN must be numeric or empty"
    End If
    adoCommand.Parameters.Append adoParameter
    
    '' AS 12/06/01 CC012 - Start
    'If eDbEngineType = SQLServer Then
        ' Pass Lender ID as binary to SQL Server.
        If Len(strLenderId) > 0 Then
            Set adoParameter = adoCommand.CreateParameter("@p_iLenderId", adBinary, adParamInput, 16, objSqlAssist.GuidStringToByteArray(strLenderId))
        Else
            Set adoParameter = adoCommand.CreateParameter("@p_iLenderId", adBinary, adParamInput, 16, Null)
        End If
        adoParameter.Attributes = adParamNullable
        adoCommand.Parameters.Append adoParameter
    'ElseIf eDbEngineType = Oracle Then
    '    ' Oracle stored procedure still expects a string; ideally this needs to be changed to match SQL Server.
    '    Set adoParameter = adoCommand.CreateParameter("@p_iLenderId", adBSTR, adParamInput, Len(strLenderId), strLenderId)
    '    adoParameter.Attributes = adParamNullable
    '    adoCommand.Parameters.Append adoParameter
    'End If
    '' AS 12/06/01 CC012 - End

    Set adoParameter = adoCommand.CreateParameter("@p_iDiscountedProducts", adInteger, adParamInput)
    adoParameter.Attributes = adParamNullable
    adoCommand.Parameters.Append adoParameter

    If Len(strDiscountedProducts) = 0 Then
        adoParameter.Value = Null
    ElseIf IsNumeric(strDiscountedProducts) Then
        adoParameter.Value = strDiscountedProducts
    Else
        m_objErrAssist.ThrowError TypeName(Me), _
                                strFunctionName, _
                                oeInvalidParameter, _
                                "DISCOUNTEDPRODUCTS must be numeric or empty"
    End If

    Set adoParameter = adoCommand.CreateParameter("@p_iDiscountedPeriod", adBSTR, adParamInput, Len(strDiscountedPeriod), strDiscountedPeriod)
    adoParameter.Attributes = adParamNullable
    adoCommand.Parameters.Append adoParameter

    Set adoParameter = adoCommand.CreateParameter("@p_iFixedRateProducts", adInteger, adParamInput)
    adoParameter.Attributes = adParamNullable
    If Len(strFixedRateProducts) = 0 Then
        adoParameter.Value = Null
    ElseIf IsNumeric(strFixedRateProducts) Then
        adoParameter.Value = strFixedRateProducts
    Else
        m_objErrAssist.ThrowError TypeName(Me), _
                                strFunctionName, _
                                oeInvalidParameter, _
                                "FIXEDRATEPRODUCTS must be numeric or empty"
    End If
    adoCommand.Parameters.Append adoParameter

    Set adoParameter = adoCommand.CreateParameter("@p_iFixedRatePeriod", adBSTR, adParamInput, Len(strFixedRatePeriod), strFixedRatePeriod)
    adoParameter.Attributes = adParamNullable
    adoCommand.Parameters.Append adoParameter

    Set adoParameter = adoCommand.CreateParameter("@p_iStandardVariableProducts", adInteger, adParamInput)
    adoParameter.Attributes = adParamNullable
    If Len(strStandardVariableProducts) = 0 Then
        adoParameter.Value = Null
    ElseIf IsNumeric(strStandardVariableProducts) Then
        adoParameter.Value = strStandardVariableProducts
    Else
        m_objErrAssist.ThrowError TypeName(Me), _
                                strFunctionName, _
                                oeInvalidParameter, _
                                "STANDARDVARIABLEPRODUCTS must be numeric or empty"
    End If
    adoCommand.Parameters.Append adoParameter

    Set adoParameter = adoCommand.CreateParameter("@p_iCappedFlooredProducts", adInteger, adParamInput)
    adoParameter.Attributes = adParamNullable
    If Len(strCappedFlooredProducts) = 0 Then
        adoParameter.Value = Null
    ElseIf IsNumeric(strCappedFlooredProducts) Then
        adoParameter.Value = strCappedFlooredProducts
    Else
        m_objErrAssist.ThrowError TypeName(Me), _
                                strFunctionName, _
                                oeInvalidParameter, _
                                "CAPPEDFLOOREDPRODUCTS must be numeric or empty"
    End If
    adoCommand.Parameters.Append adoParameter

    Set adoParameter = adoCommand.CreateParameter("@p_iCappedFlooredPeriod", adBSTR, adParamInput, Len(strCappedFlooredPeriod), strCappedFlooredPeriod)
    adoParameter.Attributes = adParamNullable
    adoCommand.Parameters.Append adoParameter

    Set adoParameter = adoCommand.CreateParameter("@p_iCashbackProducts", adInteger, adParamInput)
    adoParameter.Attributes = adParamNullable
    If Len(strCashbackProducts) = 0 Then
        adoParameter.Value = Null
    ElseIf IsNumeric(strCashbackProducts) Then
        adoParameter.Value = strCashbackProducts
    Else
        m_objErrAssist.ThrowError TypeName(Me), _
                                strFunctionName, _
                                oeInvalidParameter, _
                                "CASHBACKPRODUCTS must be numeric or empty"
    End If
    adoCommand.Parameters.Append adoParameter

    'EP2_9 GHun
    'Set adoParameter = adoCommand.CreateParameter("@p_iFlexibleProducts", adInteger, adParamInput)
    'adoParameter.Attributes = adParamNullable
    'If Len(strFlexibleProducts) = 0 Then
    '    adoParameter.Value = Null
    'ElseIf IsNumeric(strFlexibleProducts) Then
    '    adoParameter.Value = strFlexibleProducts
    'Else
    '    m_objErrAssist.ThrowError TypeName(Me), _
    '                            strFunctionName, _
    '                            oeInvalidParameter, _
    '                            "FLEXIBLEPRODUCTS must be numeric or empty"
    'End If
    'adoCommand.Parameters.Append adoParameter
    'EP2_9

    Set adoParameter = adoCommand.CreateParameter("@p_iImpairedCreditRating", adInteger, adParamInput)
    adoParameter.Attributes = adParamNullable
    If Len(strImpairedCreditRating) = 0 Then
        adoParameter.Value = Null
    ElseIf IsNumeric(strImpairedCreditRating) Then
        adoParameter.Value = strImpairedCreditRating
    Else
        m_objErrAssist.ThrowError TypeName(Me), _
                                strFunctionName, _
                                oeInvalidParameter, _
                                "IMPAIREDCREDITRATING must be numeric or empty"
    End If
    adoCommand.Parameters.Append adoParameter

    Set adoParameter = adoCommand.CreateParameter("@p_iProductsByGroup", adInteger, adParamInput)
    adoParameter.Attributes = adParamNullable
    If Len(strProductsByGroup) = 0 Then
        adoParameter.Value = Null
    ElseIf IsNumeric(strProductsByGroup) Then
        adoParameter.Value = strProductsByGroup
    Else
        m_objErrAssist.ThrowError TypeName(Me), _
                                strFunctionName, _
                                oeInvalidParameter, _
                                "PRODUCTSBYGROUP must be numeric or empty"
    End If
    adoCommand.Parameters.Append adoParameter

    '   AW  12/06/02    BM011
    Set adoParameter = adoCommand.CreateParameter("@p_iProductGroup", adInteger, adParamInput)
    adoParameter.Attributes = adParamNullable
    If Len(strProductGroup) = 0 Then
        adoParameter.Value = Null
    ElseIf IsNumeric(strProductGroup) Then
        adoParameter.Value = strProductGroup
    Else
        m_objErrAssist.ThrowError TypeName(Me), _
                                strFunctionName, _
                                oeInvalidParameter, _
                                "PRODUCTSGROUP must be numeric or empty"
    End If
    adoCommand.Parameters.Append adoParameter
    
    adoCommand.Parameters.Append adoCommand.CreateParameter("@p_iDistributionChannelId", adBSTR, adParamInput, Len(strDistributionChannelId), strDistributionChannelId)
    
    Set adoParameter = adoCommand.CreateParameter("@p_iAllProductsWithChecks", adInteger, adParamInput)
    adoParameter.Attributes = adParamNullable
    If Len(strAllProductsWithChecks) = 0 Then
        adoParameter.Value = Null
    ElseIf IsNumeric(strAllProductsWithChecks) Then
        adoParameter.Value = strAllProductsWithChecks
    Else
        m_objErrAssist.ThrowError TypeName(Me), _
                            strFunctionName, _
                            oeInvalidParameter, _
                            "ALLPRODUCTSWITHCHECKS must be numeric or empty"
    End If
    adoCommand.Parameters.Append adoParameter

    adoCommand.Parameters.Append adoCommand.CreateParameter("@p_iSearchContext", adBSTR, adParamInput, Len(strSearchContext), strSearchContext)
    
    Set adoParameter = adoCommand.CreateParameter("@p_iMemberOfStaff", adInteger, adParamInput)
    adoParameter.Attributes = adParamNullable
    If Len(strMemberOfStaff) = 0 Then
        adoParameter.Value = Null
    ElseIf IsNumeric(strMemberOfStaff) Then
        adoParameter.Value = strMemberOfStaff
    Else
        m_objErrAssist.ThrowError TypeName(Me), _
                            strFunctionName, _
                            oeInvalidParameter, _
                            "MEMBEROFSTAFF must be numeric or empty"
    End If
    adoCommand.Parameters.Append adoParameter

    Set adoParameter = adoCommand.CreateParameter("@p_iExistingCustomer", adInteger, adParamInput)
    adoParameter.Attributes = adParamNullable
    If Len(strExistingCustomer) = 0 Then
        adoParameter.Value = Null
    ElseIf IsNumeric(strExistingCustomer) Then
        adoParameter.Value = strExistingCustomer
    Else
        m_objErrAssist.ThrowError TypeName(Me), _
                                strFunctionName, _
                                oeInvalidParameter, _
                                "EXISTINGCUSTOMER must be numeric or empty"
    End If
    adoCommand.Parameters.Append adoParameter

    Set adoParameter = adoCommand.CreateParameter("@p_iTypeOfBuyer", adInteger, adParamInput)
    adoParameter.Attributes = adParamNullable
    If Len(strTypeOfBuyer) = 0 Then
        adoParameter.Value = Null
    ElseIf IsNumeric(strTypeOfBuyer) Then
        adoParameter.Value = strTypeOfBuyer
    Else
        m_objErrAssist.ThrowError TypeName(Me), _
                                strFunctionName, _
                                oeInvalidParameter, _
                                "TYPEOFBUYER must be numeric or empty"
    End If
    adoCommand.Parameters.Append adoParameter

    'CMWP3 - DPF - Jul 02 START: new field added into procedure - Portable Product Indicator
    Set adoParameter = adoCommand.CreateParameter("@p_iManualPortedLoanInd", adTinyInt, adParamInput)
    adoParameter.Attributes = adParamNullable
    If Len(strManualPortedLoanInd) = 0 Then
        adoParameter.Value = Null
    ElseIf IsNumeric(strManualPortedLoanInd) Then
        adoParameter.Value = strManualPortedLoanInd
    Else
        m_objErrAssist.ThrowError TypeName(Me), _
                                strFunctionName, _
                                oeInvalidParameter, _
                                "MANUALPORTEDLOANIND must be numeric or empty"
    End If
    adoCommand.Parameters.Append adoParameter
    'CMWP3 - END
    
    adoCommand.Parameters.Append adoCommand.CreateParameter("@p_iApplicationNumber", adBSTR, adParamInput, Len(strApplicationNumber), strApplicationNumber)

    'AW BMIDS00077 - MSMS Integration
    'AW 21/03/02 SYS4202
    Set adoParameter = adoCommand.CreateParameter("@p_iMortgageProductCode", adBSTR, adParamInput, Len(strMortgageProductCode), strMortgageProductCode)
    adoParameter.Attributes = adParamNullable
    adoCommand.Parameters.Append adoParameter
    'AW End
    
    Set adoParameter = adoCommand.CreateParameter("@p_iOrderBy", adBSTR, adParamInput, Len(strOrderBy), strOrderBy)
    adoParameter.Attributes = adParamNullable
    adoCommand.Parameters.Append adoParameter
    
    '++ BMIDS00246 SA {
    Set adoParameter = adoCommand.CreateParameter("@p_iEmploymentStatusList", adBSTR, adParamInput, Len(strEmploymentStatusList), strEmploymentStatusList)
    adoParameter.Attributes = adParamNullable
    'Need to check it's wither a value or NULL
    If Len(strEmploymentStatusList) = 0 Then
        adoParameter.Value = Null
    End If
    adoCommand.Parameters.Append adoParameter
    '++ BMIDS00246 }
        
    'MAR1702
    adoCommand.Parameters.Append adoCommand.CreateParameter("p_iMinAge", adInteger, adParamInput, , CInt(strMinAge))
    
    'MAR1702
    adoCommand.Parameters.Append adoCommand.CreateParameter("p_iMaxAge", adInteger, adParamInput, , CInt(strMaxAge))
    
    'EP2_9 GHun
    '@p_CheckFlexibleNonFlexible BIT,
    '@p_FlexibleInd          BIT
    Set adoParameter = adoCommand.CreateParameter("@p_NatureOfLoan", adInteger, adParamInput)
    adoParameter.Attributes = adParamNullable
    If Len(strNatureOfLoan) = 0 Then
        adoParameter.Value = Null
    ElseIf IsNumeric(strNatureOfLoan) Then
        adoParameter.Value = strNatureOfLoan
    Else
        m_objErrAssist.ThrowError TypeName(Me), _
                                strFunctionName, _
                                oeInvalidParameter, _
                                "NATUREOFLOAN must be numeric or empty"
    End If
    adoCommand.Parameters.Append adoParameter
    
    Set adoParameter = adoCommand.CreateParameter("@p_IncomeStatus", adInteger, adParamInput)
    adoParameter.Attributes = adParamNullable
    If Len(strIncomeStatus) = 0 Then
        adoParameter.Value = Null
    ElseIf IsNumeric(strIncomeStatus) Then
        adoParameter.Value = strIncomeStatus
    Else
        m_objErrAssist.ThrowError TypeName(Me), _
                                strFunctionName, _
                                oeInvalidParameter, _
                                "INCOMESTATUS must be numeric or empty"
    End If
    adoCommand.Parameters.Append adoParameter
    
    Set adoParameter = adoCommand.CreateParameter("@p_GuarantorInd", adBoolean, adParamInput)
    adoParameter.Attributes = adParamNullable
    If Len(strGuarantorInd) = 0 Then
        adoParameter.Value = Null
    ElseIf IsNumeric(strGuarantorInd) Then
        adoParameter.Value = strGuarantorInd
    Else
        m_objErrAssist.ThrowError TypeName(Me), _
                                strFunctionName, _
                                oeInvalidParameter, _
                                "GUARANTORIND must be numeric or empty"
    End If
    adoCommand.Parameters.Append adoParameter
    
    Set adoParameter = adoCommand.CreateParameter("@p_CheckFlexibleNonFlexible", adBoolean, adParamInput)
    adoParameter.Attributes = adParamNullable
    If Len(strFlexNonFlex) = 0 Then
        adoParameter.Value = Null
    ElseIf IsNumeric(strFlexNonFlex) Then
        adoParameter.Value = strFlexNonFlex
    Else
        m_objErrAssist.ThrowError TypeName(Me), _
                                strFunctionName, _
                                oeInvalidParameter, _
                                "FLEXIBLENONFLEXIBLEPRODUCTS must be numeric or empty"
    End If
    adoCommand.Parameters.Append adoParameter
    
    Set adoParameter = adoCommand.CreateParameter("@p_FlexibleInd", adBoolean, adParamInput)
    adoParameter.Attributes = adParamNullable
    If Len(strFlexibleInd) = 0 Then
        adoParameter.Value = Null
    ElseIf IsNumeric(strFlexibleInd) Then
        adoParameter.Value = strFlexibleInd
    Else
        m_objErrAssist.ThrowError TypeName(Me), _
                                strFunctionName, _
                                oeInvalidParameter, _
                                "FLEXIBLEIND must be numeric or empty"
    End If
    adoCommand.Parameters.Append adoParameter
    'EP2_9 End
    
    'SR EP2_76 : new inputs coresponding to ProductExclusivity
    If Len(strPrincipalFirmId) = 0 Then  '@p_iPrincipalFirmId  - NVARCHAR(24)
        Set adoParameter = adoCommand.CreateParameter("@p_iPrincipalFirmId", adBSTR, adParamInput)
        adoParameter.Attributes = adParamNullable
        adoParameter.Value = Null
        adoCommand.Parameters.Append adoParameter
    Else
        adoCommand.Parameters.Append adoCommand.CreateParameter("@p_iPrincipalFirmId", adBSTR, adParamInput, Len(strPrincipalFirmId), strPrincipalFirmId)
    End If
    
    If Len(strARFirmId) = 0 Then  '@p_iARFirmId  - NVARCHAR(24)
        Set adoParameter = adoCommand.CreateParameter("@p_iARFirmId", adBSTR, adParamInput)
        adoParameter.Attributes = adParamNullable
        adoParameter.Value = Null
        adoCommand.Parameters.Append adoParameter
    Else
        adoCommand.Parameters.Append adoCommand.CreateParameter("@p_iARFirmId", adBSTR, adParamInput, Len(strARFirmId), strARFirmId)
    End If
    
    If Len(strMortgageClubId) = 0 Then  '@p_iMortgageClubId  - NVARCHAR(24)
        Set adoParameter = adoCommand.CreateParameter("@p_iMortgageClubId", adBSTR, adParamInput)
        adoParameter.Attributes = adParamNullable
        adoParameter.Value = Null
        adoCommand.Parameters.Append adoParameter
    Else
        adoCommand.Parameters.Append adoCommand.CreateParameter("@p_iMortgageClubId", adBSTR, adParamInput, Len(strMortgageClubId), strMortgageClubId)
    End If
    
    If Len(strPackagerId) = 0 Then  '@p_iPackagerId  - NVARCHAR(24)
        Set adoParameter = adoCommand.CreateParameter("@p_iPackagerId", adBSTR, adParamInput)
        adoParameter.Attributes = adParamNullable
        adoParameter.Value = Null
        adoCommand.Parameters.Append adoParameter
    Else
        adoCommand.Parameters.Append adoCommand.CreateParameter("@p_iPackagerId", adBSTR, adParamInput, Len(strPackagerId), strPackagerId)
    End If
    
    If Len(strPackagerAssociationId) = 0 Then  '@p_iPackagerAssociationId  - NVARCHAR(24)
        Set adoParameter = adoCommand.CreateParameter("@p_iPackagerAssociationId", adBSTR, adParamInput)
        adoParameter.Attributes = adParamNullable
        adoParameter.Value = Null
        adoCommand.Parameters.Append adoParameter
    Else
        adoCommand.Parameters.Append adoCommand.CreateParameter("@p_iPackagerAssociationId", adBSTR, adParamInput, Len(strPackagerAssociationId), strPackagerAssociationId)
    End If
    'SR EP2_76 - End
    
    With rstThisRecSet
        Set .Source = adoCommand
        .CursorType = adOpenStatic
        .CursorLocation = adUseClient
        .LockType = adLockReadOnly
        .Open
    End With
    
    If Not objADOAssist.GetValidRecordset(rstThisRecSet) Then
        ' raise application error to be interpreted by calling object
        m_objErrAssist.RaiseError TypeName(Me), strFunctionName, omiga4RecordNotFound
    End If

    Set xmlOut = New FreeThreadedDOMDocument40
    xmlOut.validateOnParse = False
    xmlOut.setProperty "NewParser", True
    Set xmlListElem = xmlOut.createElement("MORTGAGEPRODUCTLIST")
    xmlOut.appendChild xmlListElem
    'xmlListElem.setAttribute "TOTAL", adoTotalParam.Value
    
    Set objIomMpClassDef = New omMPClassDef
    Set xmlClassDefDoc = objIomMpClassDef.LoadMortgageProductSelectionData()
        
    Dim xmlTest As IXMLDOMNode
    Dim blnFirstRow As Boolean
    blnFirstRow = True
    
    While Not rstThisRecSet.EOF
        If blnFirstRow Then
            xmlListElem.setAttribute "TOTAL", rstThisRecSet.Fields("MaxRecords").Value
            blnFirstRow = False
        End If
        Set xmlTest = m_objIDOAssist.GetXMLFromRecordSet(rstThisRecSet, xmlClassDefDoc, xmlListElem)
        rstThisRecSet.MoveNext
    Wend
    
    rstThisRecSet.Close
    adoConnection.Close
    
    Set IMortgageProductDO_FindProductList = xmlListElem.cloneNode(True)
    
    Set xmlOut = Nothing
    Set objSqlAssist = Nothing
    Set xmlListElem = Nothing
    Set xmlClassDefDoc = Nothing
    Set adoConnection = Nothing
    Set adoCommand = Nothing
    Set rstThisRecSet = Nothing
    Set adoParameter = Nothing
    Set objIomMpClassDef = Nothing
    Set objADOAssist = Nothing
    Set objGlobalParamDO = Nothing
        
    If Not m_objContext Is Nothing Then
        m_objContext.SetComplete
    End If

    Exit Function

FindProductListVbErr:
    
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

    Set xmlOut = Nothing
    Set objSqlAssist = Nothing
    Set xmlListElem = Nothing
    Set xmlClassDefDoc = Nothing
    Set adoConnection = Nothing
    Set adoCommand = Nothing
    Set rstThisRecSet = Nothing
    Set adoParameter = Nothing
    Set objIomMpClassDef = Nothing
    Set objADOAssist = Nothing
    Set objGlobalParamDO = Nothing
    
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

Private Function IMortgageProductDO_GetValidActiveProduct(ByVal vxmlTableElement As IXMLDOMElement) As Boolean

' header ----------------------------------------------------------------------------------
' description:
'   Get the data for a single instance of the persistant data associated with
'   this data object
' pass:
'   vxmlxmlTableElement  xml element containing the request
'   vxmlInNode      xml node to which data should be attached
' return:
'                   xml node that has been attached
'------------------------------------------------------------------------------------------
On Error GoTo GetValidActiveProductVbErr

    Const strFunctionName As String = "IMortgageProductDO_GetValidActiveProduct"
    
    'Dim objDOAssist As DOAssist
    Dim objADOAssist As ADOAssist
    Dim objSqlAssist As SQLAssist
    Dim xmlClassDefDoc As FreeThreadedDOMDocument40
            
    If m_objContext Is Nothing Then
        'Set objDOAssist = New DOAssist
        Set objADOAssist = New ADOAssist
    Else
        'Set objDOAssist = m_objContext.CreateInstance(gstrBASE_COMPONENT & ".DOAssist")
        Set objADOAssist = m_objContext.CreateInstance(gstrBASE_COMPONENT & ".ADOAssist")
    End If
    
    Dim strMortgageProductCode As String
    Dim strMortgageProductStartDate As String
    Dim strStaffMortgageProduct As String
    ''BMIDS00952 SA 15/11/2002 Self Employed no longer available as product option so remove it here.
    'Dim strSelfEmployed As String
    Dim strApplicationNumber As String
    Dim strLTV As String                'BMIDS00515 SA 24/10/02
    
    strMortgageProductCode = m_objXmlAssist.GetTagValue(vxmlTableElement, "MORTGAGEPRODUCTCODE")
    strMortgageProductStartDate = m_objXmlAssist.GetTagValue(vxmlTableElement, "STARTDATE")
    strStaffMortgageProduct = m_objXmlAssist.GetTagValue(vxmlTableElement, "ISMEMBEROFSTAFF")
    'BMIDS00952 SA 15/11/2002 Self Employed no longer available as product option so remove it here.
    'strSelfEmployed = m_objXmlAssist.GetTagValue(vxmlTableElement, "ISSELFEMPLOYED")
    strApplicationNumber = m_objXmlAssist.GetTagValue(vxmlTableElement, "APPLICATIONNUMBER")
    strLTV = m_objXmlAssist.GetTagValue(vxmlTableElement, "LTV")  'BMIDS00515 SA 24/10/02
    
    If Len(strMortgageProductCode) = 0 Or Len(strMortgageProductStartDate) = 0 Or _
        Len(strApplicationNumber) = 0 Then
    'BMIDS00952 SA 15/11/2002 Self Employed no longer available as product option so remove it here.
        m_objErrAssist.RaiseError TypeName(Me), _
                        strFunctionName, _
                        omiga4MissingPrimaryTag, _
                        "MORTGAGEPRODUCTCODE =" & strMortgageProductCode & _
                        ",STARTDATE =" & strMortgageProductStartDate & _
                        ",ISMEMBEROFSTAFF = " & strStaffMortgageProduct & _
                        ",APPLICATIONNUMBER =" & strApplicationNumber & "Expected"
    
    End If
    'Set GetValidActiveProduct = objDOAssist.GetDataEx(vxmlTableElement, LoadData, vxmlInNode)
    
    Set objSqlAssist = New SQLAssist
    
    ' APS 09/05/2000 - Changed the SQL statement only needs to check the RESERVEDPRODUCT
    ' table if the product has been withdrawn
    Dim strSQL As String
    Dim strSelect As String
    strSQL = " MORTGAGEPRODUCT.MORTGAGEPRODUCTCODE = " & _
            objSqlAssist.FormatString(strMortgageProductCode) & _
            " AND MORTGAGEPRODUCT.STARTDATE = " & _
            objSqlAssist.FormatDateString(strMortgageProductStartDate) & _
            " AND ( " & objSqlAssist.GetSystemDate & " >= MORTGAGEPRODUCT.STARTDATE )" & _
            " AND ( " & objSqlAssist.GetSystemDate & " < ENDDATE ) "
            
    If (StrComp(strStaffMortgageProduct, "0", vbTextCompare) = 0) Then
        strSQL = strSQL & " AND MORTGAGEPRODUCT.STAFFMORTGAGEPRODUCT = " & strStaffMortgageProduct
    End If
    'BMIDS00952 SA 15/11/2002 Self Employed no longer available as product option so remove it here.
'    If (StrComp(strSelfEmployed, "1", vbTextCompare) = 0) Then
'        strSQL = strSQL & " AND MORTGAGEPRODUCT.AVAILABLEFORSELFEMPLOYED = " & strSelfEmployed
'    End If
    'BMIDS00515 SA 24/10/02 {
    If Len(strLTV) > 0 Then             ' SA BMIDS00515 29/10/02 check to see if LTV is actually passed in!
        If (StrComp(strLTV, "0", vbTextCompare) <> 0) Then
            strSQL = strSQL & " AND MORTGAGEPRODUCT.MINIMUMLTV <= " & strLTV & " AND MORTGAGEPRODUCT.MAXIMUMLTV >= " & strLTV
        End If
    End If
    'BMIDS00515 }
                    
    strSQL = strSQL & " AND (( WITHDRAWNDATE IS NULL OR WITHDRAWNDATE > " & objSqlAssist.GetSystemDate & ") OR" & _
            " ( WITHDRAWNDATE <= " & objSqlAssist.GetSystemDate & " AND MORTGAGEPRODUCT.MORTGAGEPRODUCTCODE IN " & _
            " (SELECT RESERVEDPRODUCT.MORTGAGEPRODUCTCODE FROM RESERVEDPRODUCT WHERE " & _
            " RESERVEDPRODUCT.MORTGAGEPRODUCTCODE = MORTGAGEPRODUCT.MORTGAGEPRODUCTCODE" & _
            " AND RESERVEDPRODUCT.STARTDATE = MORTGAGEPRODUCT.STARTDATE" & _
            " AND RESERVEDPRODUCT.APPLICATIONNUMBER = " & objSqlAssist.FormatString(strApplicationNumber) & ")))"
    strSelect = "SELECT MORTGAGEPRODUCT.MORTGAGEPRODUCTCODE FROM " & _
            "MORTGAGEPRODUCT" & " Where " & strSQL                    ' PSC 10/07/00 SYS1193
    
    Dim cmd As ADODB.Command
    Set cmd = New ADODB.Command
    cmd.ActiveConnection = objADOAssist.GetConnStr
    cmd.CommandText = strSelect
    
    Dim rstThisRecSet As ADODB.Recordset

    Set rstThisRecSet = cmd.Execute(1, , adCmdText)
            
'    Set xmlClassDefDoc = LoadMortgageProductDataEx()
        
    If Not rstThisRecSet.EOF Then
        ' convert record set to xml
        IMortgageProductDO_GetValidActiveProduct = True
'        Set GetValidActiveProduct = objDOAssist.GetXMLFromRecordSetEx(rstThisRecSet, xmlClassDefDoc, vxmlInNode)
        'if record not found then return false else return true
    Else
        ' raise application error to be interpreted by calling object
        'objErrAssist.RaiseError TypeName(Me), strFunctionName, omiga4RecordNotFound
        IMortgageProductDO_GetValidActiveProduct = False
    End If
    rstThisRecSet.Close

    Set rstThisRecSet = Nothing
    Set cmd = Nothing
    Set objADOAssist = Nothing

    Set xmlClassDefDoc = Nothing
    'Set objDOAssist = Nothing
    
    If Not m_objContext Is Nothing Then
        m_objContext.SetComplete
    End If
    
    Exit Function

GetValidActiveProductVbErr:
        
    If Err.Number = omiga4RecordNotFound Then
        Resume Next
    End If
        
    ' Close the recordset if still open
    If Not rstThisRecSet Is Nothing Then
        If rstThisRecSet.State = adStateOpen Then
            rstThisRecSet.Close
        End If
    End If

    Set rstThisRecSet = Nothing
    Set cmd = Nothing
    Set xmlClassDefDoc = Nothing
    'Set objDOAssist = Nothing
    
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

Private Function IMortgageProductDO_GetInterestRateType(ByVal vxmlRequest As IXMLDOMElement) As IXMLDOMNode
' header ----------------------------------------------------------------------------------
' description:
'   Get the data for all instances of the persistant data associated with
'   this data object for the values supplied
' pass:
'   vstrXMLRequest  xml Request data stream containing data to which identifies
'                   the instance(s) of the persistant data to be retrieved
' return:
'   FindList        string containing XML data stream representation of
'                   data retrieved
' Raise Errors:
'   if record not found, raise omiga4RecordNotFound
'------------------------------------------------------------------------------------------
On Error GoTo IMortgageProductDO_GetInterestRateTypevbErr

    Const strFunctionName As String = "IMortgageProductDO_GetInterestRateType"

    Dim objSqlAssist As SQLAssist
    Dim objConvertAssist As ConvertAssist
    Dim objXmlInterestRateNode As IXMLDOMNode
    
    Dim objXmlOut As FreeThreadedDOMDocument40
    Dim objXMLInt As IXMLDOMElement

    Set objSqlAssist = New SQLAssist
    Set objConvertAssist = New ConvertAssist
    Dim xmlNode As IXMLDOMNode
    Dim xmlNodeList As IXMLDOMNodeList
    Dim xmlLoadInterestRateTypeData As FreeThreadedDOMDocument40
        
    Dim IomMPClassDef As IomMPClassDef
    
    Set IomMPClassDef = New omMPClassDef
    
    Set xmlLoadInterestRateTypeData = IomMPClassDef.LoadInterestRateTypeData
    Set xmlNode = vxmlRequest.cloneNode(True)

    Set objXmlInterestRateNode = m_objIDOAssist.FindList(xmlNode, xmlLoadInterestRateTypeData)
    Set objXmlOut = New FreeThreadedDOMDocument40
    objXmlOut.validateOnParse = False
    objXmlOut.setProperty "NewParser", True
    Set objXMLInt = objXmlOut.appendChild(objXmlInterestRateNode)
    Set xmlNodeList = objXMLInt.getElementsByTagName("INTERESTRATETYPE")
    For Each xmlNode In xmlNodeList
      If (m_objXmlAssist.GetElementText(xmlNode, "INTERESTRATETYPESEQUENCENUMBER") = "1") Then
        Set IMortgageProductDO_GetInterestRateType = xmlNode
      End If
    Next
    Set xmlNode = Nothing
    Set xmlNodeList = Nothing

    If Not m_objContext Is Nothing Then
        m_objContext.SetComplete
    End If
    
    Exit Function

IMortgageProductDO_GetInterestRateTypevbErr:
    
    Set xmlNode = Nothing
    Set xmlNodeList = Nothing
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

' PSC 18/06/2006 MAR1687 - Start
Private Function IMortgageProductDO_GetFullProductDetails(ByVal vxmlRequest As IXMLDOMElement) As IXMLDOMNode

    On Error GoTo GetFullProductDetailsErr
    
    Const cstrFunctionName  As String = "IMortgageProductDO_GetFullProductDetails"
    
    Dim adoConnection       As ADODB.Connection
    Dim adoCommand          As ADODB.Command
    Dim adoStream           As ADODB.Stream
    Dim xmlDoc              As FreeThreadedDOMDocument40
    Dim xmlMortgageProduct  As IXMLDOMNode
    Dim xmlProductDetails   As IXMLDOMNode
    Dim xmlBandList         As IXMLDOMNode
    Dim objConvertAssist    As ConvertAssist
    Dim objIADOAssist       As IADOAssist

    Dim strProductCode As String
    Dim dteStartDate As Date
    Dim lngAmountRequested As Long
    Dim lngPurchasePrice As Long
    Dim dblLTV As Double
    Dim intTypeOfApplication As Integer
    Dim intLocation As Integer
    Dim intValuationType As Integer
    Dim intLegalFeeType As Integer
    Dim blnPortedLoan As Boolean
    Dim strPortingStartDate As String
    'EP2_9 GHun
    Dim strApplicationNumber As String
    Dim dteDateAppCreated As Date
    'EP2_9 End
    
    Dim dblOriginalLTV As Double     ' PSC 19/02/2007

    Dim adoParam As ADODB.Parameter 'EP2_126 GHun
    
    Set objConvertAssist = New ConvertAssist
    Set objIADOAssist = New ADOAssist
    
    Set adoConnection = New ADODB.Connection
    Set adoCommand = New ADODB.Command
    
    If vxmlRequest.nodeName = "MORTGAGEPRODUCT" Then
        Set xmlMortgageProduct = vxmlRequest
    Else
        Set xmlMortgageProduct = m_objXmlAssist.GetMandatoryNode(vxmlRequest, "MORTGAGEPRODUCT")
    End If
     
    strProductCode = m_objXmlAssist.GetMandatoryElementText(xmlMortgageProduct, "MORTGAGEPRODUCTCODE")
    dteStartDate = objConvertAssist.CSafeDate(m_objXmlAssist.GetMandatoryElementText(xmlMortgageProduct, "STARTDATE"))
    lngAmountRequested = objConvertAssist.CSafeLng(m_objXmlAssist.GetMandatoryElementText(xmlMortgageProduct, "AMOUNTREQUESTED"))
    lngPurchasePrice = objConvertAssist.CSafeLng(m_objXmlAssist.GetElementText(xmlMortgageProduct, "PURCHASEPRICE"))
    dblLTV = objConvertAssist.CSafeDbl(m_objXmlAssist.GetMandatoryElementText(xmlMortgageProduct, "LTV"))
    intTypeOfApplication = objConvertAssist.CSafeInt(m_objXmlAssist.GetMandatoryElementText(xmlMortgageProduct, "TYPEOFAPPLICATION"))
    intLocation = objConvertAssist.CSafeInt(m_objXmlAssist.GetMandatoryElementText(xmlMortgageProduct, "LOCATION"))
    intValuationType = objConvertAssist.CSafeInt(m_objXmlAssist.GetMandatoryElementText(xmlMortgageProduct, "TYPEOFVALUATION"))
    intLegalFeeType = objConvertAssist.CSafeInt(m_objXmlAssist.GetMandatoryElementText(xmlMortgageProduct, "FEETYPE"))
    'EP2_9 GHun
    strApplicationNumber = m_objXmlAssist.GetElementText(xmlMortgageProduct, "APPLICATIONNUMBER")   'EP2_126 GHun
    dteDateAppCreated = objConvertAssist.CSafeDate(m_objXmlAssist.GetMandatoryElementText(xmlMortgageProduct, "APPLICATIONDATE"))
    'EP2_9 End
    
    ' PSC 19/02/2007 EP2_1488 - Start
    dblOriginalLTV = objConvertAssist.CSafeDbl(m_objXmlAssist.GetElementText(xmlMortgageProduct, "ORIGINALLTV"))
    
    ' Original LTV overrides current LTV
    If dblOriginalLTV > 0 Then
        dblLTV = dblOriginalLTV
    End If
    ' PSC 19/02/2007 EP2_1488 - End
    
    With adoConnection
        .ConnectionString = objIADOAssist.GetConnStr
        .CursorLocation = adUseClient
        .Open
    End With
    With adoCommand
        .CommandText = "USP_GETFULLPRODUCTDETAILS"
        .CommandType = adCmdStoredProc
        .Parameters.Append .CreateParameter("@ProductCode", adBSTR, adParamInput, Len(strProductCode), strProductCode)
        .Parameters.Append .CreateParameter("@StartDate", adDBTimeStamp, adParamInput, , dteStartDate)
        .Parameters.Append .CreateParameter("@AmountRequested", adInteger, adParamInput, , lngAmountRequested)
        .Parameters.Append .CreateParameter("@PurchasePrice", adInteger, adParamInput, , lngPurchasePrice)
        .Parameters.Append .CreateParameter("@LTV", adDouble, adParamInput, , dblLTV)
        .Parameters.Append .CreateParameter("@TypeOfApplication", adInteger, adParamInput, , intTypeOfApplication)
        .Parameters.Append .CreateParameter("@Location", adInteger, adParamInput, , intLocation)
        .Parameters.Append .CreateParameter("@ValuationType", adInteger, adParamInput, , intValuationType)
        .Parameters.Append .CreateParameter("@LegalFeeType", adInteger, adParamInput, , intLegalFeeType)
        'EP2_9 GHun
        'EP2_126 GHun
        Set adoParam = .CreateParameter("@ApplicationNumber", adBSTR, adParamInput, Len(strApplicationNumber), strApplicationNumber)
        If Len(strApplicationNumber) = 0 Then
            adoParam.Attributes = adParamNullable
            adoParam.Value = Null
        End If
        .Parameters.Append adoParam
        'EP2_126 End
        .Parameters.Append .CreateParameter("@DateAppCreated", adDBTimeStamp, adParamInput, , dteDateAppCreated)
        'EP2_9 End
    End With
    Set adoStream = New ADODB.Stream
    adoStream.Open
        
    With adoCommand
        .ActiveConnection = adoConnection
        .Properties("Output Stream") = adoStream
        .Properties("XML Root") = "MORTGAGEPRODUCTDETAILS"
        .Execute , , adExecuteStream
        Set .ActiveConnection = Nothing
    End With
        
    Set xmlDoc = New FreeThreadedDOMDocument40
    xmlDoc.setProperty "NewParser", True
    xmlDoc.validateOnParse = False
    If xmlDoc.loadXML(adoStream.ReadText) Then
        Set xmlProductDetails = xmlDoc.selectSingleNode("MORTGAGEPRODUCTDETAILS")
        
        blnPortedLoan = (m_objXmlAssist.GetElementText(xmlMortgageProduct, "PORTINGINDICATOR") = "1")
        
        If blnPortedLoan Then
            Set xmlBandList = m_objXmlAssist.GetMandatoryNode(xmlProductDetails, "INTERESTRATETYPELIST")
            strPortingStartDate = m_objXmlAssist.GetElementText(xmlMortgageProduct, "PORTEDPRODUCTSTARTDATE")
            ProcessPortedLoanBands xmlBandList, strPortingStartDate
            Set xmlBandList = m_objXmlAssist.GetMandatoryNode(xmlProductDetails, "REDEMPTIONFEEBANDLIST")
            ProcessPortedLoanBands xmlBandList, strPortingStartDate
        End If
        
        
    End If
    
    Set IMortgageProductDO_GetFullProductDetails = xmlProductDetails
    
    adoConnection.Close
    
    Set adoConnection = Nothing
    Set adoConnection = Nothing
    Set adoCommand = Nothing
    Set adoStream = Nothing
    Set xmlDoc = Nothing
    Set xmlMortgageProduct = Nothing
    Set xmlProductDetails = Nothing
    Set xmlBandList = Nothing
    
    Set objConvertAssist = Nothing
    Set objIADOAssist = Nothing
    Set adoParam = Nothing  'EP2_126 GHun
    
    Exit Function
    
GetFullProductDetailsErr:
    
    If Not m_objContext Is Nothing Then
        m_objContext.SetComplete
    End If
    
    ' Close connection if not already closed
    If Not adoConnection Is Nothing Then
        If adoConnection.State = adStateOpen Then
            adoConnection.Close
        End If
    End If
    
    Set adoConnection = Nothing
    Set adoCommand = Nothing
    Set adoStream = Nothing
    Set xmlDoc = Nothing
    Set xmlMortgageProduct = Nothing
    Set xmlProductDetails = Nothing
    Set xmlBandList = Nothing

    Set objConvertAssist = Nothing
    Set objIADOAssist = Nothing
    Set adoParam = Nothing  'EP2_126 GHun
    
    Err.Raise Err.Number, Err.Source, Err.Description

End Function
' PSC 18/06/2006 MAR1687 - End

Private Sub ObjectControl_Activate()

    Set m_objContext = GetObjectContext()
    Set m_objErrAssist = New ErrAssist
    Set m_objXmlAssist = New XMLAssist
    Set m_objIDOAssist = m_objContext.CreateInstance(gstrBASE_COMPONENT & ".DOAssist")

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

Private Function IMortgageProductDO_GetMortgageProductLanguage(ByVal vxmlTableElement As MSXML2.IXMLDOMElement) As MSXML2.IXMLDOMNode
On Error GoTo GetMortgageProductLanguageVbErr

Const strFunctionName As String = "IMortgageProductDO_GetMortgageProductLanguage"
    
Dim xmlClassDefDoc As FreeThreadedDOMDocument40
Dim objIomMpClassDef As IomMPClassDef
Dim objIADOAssist As IADOAssist

'SYS3440 MDC 10/12/2001
Dim xmlNode As IXMLDOMNode
Dim objGlobalParameterDO As GlobalParameterDO
'SYS3440 End
Dim objConnection As ADODB.Connection
Dim objRs As ADODB.Recordset

Dim strSQL As String

    If m_objContext Is Nothing Then
        Set objIADOAssist = New ADOAssist
        Set objGlobalParameterDO = New GlobalParameterDO    'SYS3440 MDC 10/12/2001
    Else
        Set objIADOAssist = m_objContext.CreateInstance(gstrBASE_COMPONENT & ".ADOAssist")
        Set objGlobalParameterDO = m_objContext.CreateInstance(gstrBASE_COMPONENT & ".GlobalParameterDO") 'SYS3440 MDC 10/12/2001
    End If
    
    Set objConnection = New ADODB.Connection
    Set objRs = New ADODB.Recordset
    
    strSQL = "SELECT * FROM MORTGAGEPRODUCTLANGUAGE"
    
    objRs.CursorLocation = adUseClient
    objRs.CursorType = adOpenStatic
    objRs.LockType = adLockReadOnly
    objConnection.ConnectionString = objIADOAssist.GetConnStr
    objConnection.Open
    objRs.Open strSQL, objConnection, adOpenStatic, adLockReadOnly
    Set objRs.ActiveConnection = Nothing
    Set objConnection = Nothing
    
    Set objIomMpClassDef = New omMPClassDef
    Set xmlClassDefDoc = objIomMpClassDef.LoadMortgageProductLanguageData

    'Set xmlRecords = New FreeThreadedDOMDocument40
     Dim xmlOut As FreeThreadedDOMDocument40
    Dim xmlRecords As IXMLDOMNode
    Set xmlOut = New FreeThreadedDOMDocument40
    xmlOut.validateOnParse = False
    xmlOut.setProperty "NewParser", True
    Set xmlRecords = xmlOut.createElement("MORTGAGEPRODUCTLANGUAGE")
    
    If Not objRs.EOF Then
         Call m_objIDOAssist.GetXMLFromWholeRecordset(objRs, xmlClassDefDoc, xmlRecords)
    Else
        ' raise application error to be interpreted by calling object
        m_objErrAssist.RaiseError TypeName(Me), strFunctionName, omiga4RecordNotFound, "MortgageProductLanguage"
    End If
    
    Set IMortgageProductDO_GetMortgageProductLanguage = xmlRecords
    
    objRs.Close
    
    Set objIomMpClassDef = Nothing
    Set xmlClassDefDoc = Nothing
    Set objIADOAssist = Nothing
    Set objRs = Nothing
    'SYS3440 MDC 10/12/2001
    Set xmlNode = Nothing
    Set objGlobalParameterDO = Nothing
    'SYS3440 End
    
    If Not m_objContext Is Nothing Then
        m_objContext.SetComplete
    End If
    
    Exit Function

GetMortgageProductLanguageVbErr:
    
    Set objIomMpClassDef = Nothing
    Set xmlClassDefDoc = Nothing
    Set objConnection = Nothing
    
    'SYS3440 MDC 10/12/2001
    Set xmlNode = Nothing
    Set objGlobalParameterDO = Nothing
    'SYS3440 End
    
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

Private Function IMortgageProductDO_FindProductConditionsList(ByVal vxmlTableElement As IXMLDOMElement) As IXMLDOMNode
' header ----------------------------------------------------------------------------------
' description:
'   Get the data for a single instance of the persistant data associated with
'   this data object
' pass:
'   vxmlxmlTableElement  xml element containing the request
' return:                xml node containing retrieved data
'------------------------------------------------------------------------------------------
On Error GoTo FindProductConditionsListVbErr

    Const strFunctionName As String = "IMortgageProductDO_FindProductConditionsList"
           
    If vxmlTableElement Is Nothing Then
        m_objErrAssist.ThrowError TypeName(Me), _
                                strFunctionName, _
                                oeInvalidParameter, _
                                "Request node is nothing"
    End If
    Dim objADOAssist As ADOAssist

    If m_objContext Is Nothing Then
      Set objADOAssist = New ADOAssist
    Else
        Set objADOAssist = m_objContext.CreateInstance(gstrBASE_COMPONENT & ".ADOAssist")
    End If
    
    '------------------------------------------------------------------------------------------------------------
    'STEP 1 : Get strValueID = value id from combo group ConditionType where validation type is 'P' - Product
        'Dim colValueIDs As Collection
    Dim strValueID As String
    Dim objComboDO As ComboDO

    Set objComboDO = New ComboDO
    'Get the value id for ConditionType where the validation type is 'P' - Product Special Condition
    strValueID = objComboDO.GetFirstComboValueId("ConditionType", "P")

    'If no records found in ConditionType with validation type of 'P', raise an error.
    If Len(strValueID) = 0 Then
        m_objErrAssist.ThrowError TypeName(Me), strFunctionName, oeRecordNotFound, "Can't find any value IDs for ConditionType of 'P'"
    End If
    '------------------------------------------------------------------------------------------------------------
    'STEP 2 : Build SQL for getting ConditionReference from MortgageProductCondition = strSQL
    '           using (In) MortgageProductCode and (In) StartDate (there are potentially 1 or more of these).
    '           Return Unique rows only.    Dim xmlMortgageProductList As IXMLDOMNodeList
    Dim iIndex As Integer
    Dim iNumMP As Integer
    
    Dim xmlMPElem As IXMLDOMElement
    Dim xmlMortgageProductList As IXMLDOMNodeList
       
    Dim strMortgageProductCode As String
    Dim strStartDate As String
    Dim strWhere As String
    Dim strAnd As String
    Dim strSQL As String
    Dim strMainSQL As String
    
    strWhere = " WHERE ( "
    
    Set xmlMortgageProductList = vxmlTableElement.selectNodes(".//MORTGAGEPRODUCT")
    If xmlMortgageProductList Is Nothing Then
        'Can't find any MORTGAGEPRODUCTS - Raise Error and exit
        m_objErrAssist.ThrowError TypeName(Me), strFunctionName, oeMissingPrimaryTag, "No MORTGAGEPRODUCT nodes found."
    End If
    iNumMP = xmlMortgageProductList.length
    For iIndex = 0 To (iNumMP - 1)
        strAnd = vbNullString
        Set xmlMPElem = xmlMortgageProductList.Item(iIndex)
        'Check for presence of MORTGAGEPRODUCTCODE and STARTDATE, and that they have values
        'MORTGAGEPRODUCTCODE
        If Not (xmlMPElem.selectSingleNode(".//MORTGAGEPRODUCTCODE") Is Nothing) Then
            strMortgageProductCode = xmlMPElem.selectSingleNode(".//MORTGAGEPRODUCTCODE").Text
            If Len(strMortgageProductCode) = 0 Then
                'Raise Error - MORTGAGEPRODUCTCODE node present, but empty.
                m_objErrAssist.ThrowError TypeName(Me), strFunctionName, oeMissingPrimaryTag, "MORTGAGEPRODUCTCODE node present, but empty."
            End If
        Else
            'Raise Error - Missing MORTGAGEPRODUCTCODE node
            m_objErrAssist.ThrowError TypeName(Me), strFunctionName, oeMissingPrimaryTag, "Missing MORTGAGEPRODUCTCODE node."
        End If
        
        'STARTDATE
        If Not (xmlMPElem.selectSingleNode(".//STARTDATE") Is Nothing) Then
            strStartDate = xmlMPElem.selectSingleNode(".//STARTDATE").Text
            If Len(strStartDate) = 0 Then
                'Raise Error - STARTDATE node present, but empty.
                m_objErrAssist.ThrowError TypeName(Me), strFunctionName, oeMissingPrimaryTag, "STARTDATE node present, but empty."
            End If
        Else
            'Raise Error - Missing STARTDATE node
            m_objErrAssist.ThrowError TypeName(Me), strFunctionName, oeMissingPrimaryTag, "Missing STARTDATE node."
        End If
        strAnd = " ( MORTGAGEPRODUCTCODE = '" & strMortgageProductCode & "' AND STARTDATE = '" & strStartDate & "')"
        If iIndex <> (iNumMP - 1) Then 'tack on an 'OR' if not the last one in the list
            strAnd = strAnd & " OR "
        End If
        strWhere = strWhere & strAnd
    Next iIndex
    strWhere = strWhere & " )"
    '------------------------------------------------------------------------------------------------------------
    'STEP 3 : Use this result to get all the rows from the Conditions table which match the Conditionreferences returned above
    '           and where ConditionType = strValueID.

    strSQL = "SELECT DISTINCT(MORTGAGEPRODUCTCONDITION.CONDITIONREFERENCE) FROM MORTGAGEPRODUCTCONDITION "

    strSQL = strSQL & strWhere
    'strSQL will contain SQL in this structure :
    
    '    SELECT DISTINCT(MORTGAGEPRODUCTCONDITION.CONDITIONREFERENCE)
    '    FROM
    '    MORTGAGEPRODUCTCONDITION
    '    WHERE
    '    (
    '    (MORTGAGEPRODUCTCODE = 'swp2' and STARTDATE = '01/01/2001')
    '    or
    '    (MORTGAGEPRODUCTCODE = '12345' and STARTDATE = '13/06/2002 09:00:00')
    '    )
    
    strMainSQL = "SELECT C.* FROM ( " & strSQL & " ) COND LEFT OUTER JOIN CONDITIONS C ON C.CONDITIONREFERENCE = COND.CONDITIONREFERENCE WHERE C.CONDITIONTYPE = " & strValueID
    
    Debug.Print "strMainSQL = " & strMainSQL
    ' strMainSQL wil contain SQL in this structure
        'SELECT C.* FROM
        '(
        'SELECT DISTINCT(MORTGAGEPRODUCTCONDITION.CONDITIONREFERENCE)
        'From MORTGAGEPRODUCTCONDITION
        'Where
        '(
        '( MORTGAGEPRODUCTCODE = 'swp2' AND STARTDATE = '01/01/2001')
        'OR
        '( MORTGAGEPRODUCTCODE = '12345' AND STARTDATE = '13/06/2002 09:00:00')
        'OR
        '( MORTGAGEPRODUCTCODE = 'somecode' AND STARTDATE = '09/09/2001')
        ')
        '
        ') COND LEFT OUTER JOIN CONDITIONS C
        '
        '
        'ON C.CONDITIONREFERENCE = COND.CONDITIONREFERENCE WHERE C.CONDITIONTYPE = 30
    '------------------------------------------------------------------------------------------------------------
        
    'STEP 4 Execute the SQL
    Dim objComm As ADODB.Command
    Dim objRs As ADODB.Recordset
    Dim xmlOut As FreeThreadedDOMDocument40
    Dim xmlListElem As IXMLDOMElement
    Dim xmlTemp As IXMLDOMElement
    Dim objIomMpClassDef As IomMPClassDef
    Dim xmlClassDefDoc As FreeThreadedDOMDocument40
    
    Set objComm = New ADODB.Command
    Set objRs = New ADODB.Recordset
    Set xmlOut = New FreeThreadedDOMDocument40
    xmlOut.validateOnParse = False
    xmlOut.setProperty "NewParser", True
    
    Set xmlListElem = xmlOut.createElement("CONDITIONSLIST")
    xmlOut.appendChild xmlListElem
    
    With objComm
        .ActiveConnection = objADOAssist.GetConnStr
        .CommandType = adCmdText
        .CommandText = strMainSQL
        Set objRs = .Execute()
    End With
    
    If objRs.EOF Then
        ' raise application error to be interpreted by calling object
        m_objErrAssist.RaiseError TypeName(Me), strFunctionName, omiga4RecordNotFound, "Conditions"
    End If
   
    Set objIomMpClassDef = New omMPClassDef
    Set xmlClassDefDoc = objIomMpClassDef.LoadConditionsData()
    
    While Not objRs.EOF
        Set xmlTemp = m_objIDOAssist.GetXMLFromRecordSet(objRs, xmlClassDefDoc, xmlListElem)
        objRs.MoveNext
    Wend

    objRs.Close
    Set objComm = Nothing
    
'------------------------------------------------------------------------------------------------------------
            
    Set IMortgageProductDO_FindProductConditionsList = xmlListElem.cloneNode(True)

    If Not m_objContext Is Nothing Then
        m_objContext.SetComplete
    End If

    Exit Function

FindProductConditionsListVbErr:
    
    Set objADOAssist = Nothing
    Set objComboDO = Nothing
    Set xmlMPElem = Nothing
    Set xmlMortgageProductList = Nothing
    Set objComm = Nothing
    Set objRs = Nothing
    Set xmlOut = Nothing
    Set xmlListElem = Nothing
    Set xmlTemp = Nothing
    Set objIomMpClassDef = Nothing
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
   
End Function

'MAR1702
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
    
    Dim objErrAssist As ErrAssist
    Dim intYears As Integer
    Dim dteNow As Date
    Dim dteDOB As Date
    
    If (Len(vstrDateOfBirth) < Len("dd/mm/yyyy")) Or _
       (Not IsDate(vstrDateOfBirth)) Then
        Set objErrAssist = New ErrAssist
        objErrAssist.ThrowError TypeName(Me), strFunctionName, oeInvalidDateTimeFormat, _
                        " Expected Date Of Birth in format DD/MM/YYYY, but received: " & vstrDateOfBirth
    End If
        
    dteDOB = CDate(vstrDateOfBirth)
    dteNow = Now
    intYears = Year(dteNow) - Year(dteDOB)
    
    If DateAdd("yyyy", intYears, dteDOB) > dteNow Then
        intYears = intYears - 1
    End If
    
    GetAge = intYears
    
    Set objErrAssist = Nothing
    
    Exit Function
    
GetAgeVbErr:
    
    Set objErrAssist = Nothing

    '   re-raise error for business object to interpret as appropriate
    Err.Raise Err.Number, Err.Source, Err.Description

End Function

' PSC 18/06/2006 MAR1687 - Start
Private Sub ProcessPortedLoanBands(ByVal vxmlBandList As IXMLDOMNode, ByVal vstrPortingStartDate As String)
On Error GoTo ProcessPortedLoanBandsErr

    Const strFunctionName As String = "ProcessPortedLoanInterestRates"

    Dim xmlNode                 As IXMLDOMNode
    Dim xmlPeriod               As IXMLDOMNode
    
    Dim blnFound                As Boolean
    Dim strDate                 As String
    Dim dtmStartDate            As Date
    Dim dtmPeriodEndDate        As Date
    Dim intPeriodDifference     As Integer
    Dim intExpiredPeriod        As Integer
    Dim intTotalExpiredPeriod   As Integer
    Dim intRemainingPeriod      As Integer
    
    blnFound = False
    
    strDate = vstrPortingStartDate
      
    If IsDate(strDate) Then
        dtmStartDate = CDate(strDate)
        intPeriodDifference = DateDiff("m", dtmStartDate, Date)
        'Always round up
        'If there are any days difference remaining after taking the months into account,
        'then increment the period by 1
        If DateDiff("d", DateAdd("m", intPeriodDifference, dtmStartDate), Date) > 0 Then
            intPeriodDifference = intPeriodDifference + 1
        End If

        intTotalExpiredPeriod = 0
    
        Set xmlNode = vxmlBandList.firstChild
    
        'Loop through the interest rate bands until the current band is found
        Do While Not (blnFound Or (xmlNode Is Nothing))

            If xmlNode.nodeName = "INTERESTRATETYPE" Then
                strDate = m_objXmlAssist.GetElementText(xmlNode, "INTERESTRATEENDDATE")
            Else
                strDate = m_objXmlAssist.GetElementText(xmlNode, "PERIODENDDATE")
            End If

            If IsDate(strDate) Then
                'PeriodEndDate
                dtmPeriodEndDate = CDate(strDate)
                If Date < dtmPeriodEndDate Then
                    blnFound = True
                Else
                    intExpiredPeriod = DateDiff("m", dtmStartDate, dtmPeriodEndDate)
                    intTotalExpiredPeriod = intTotalExpiredPeriod + intExpiredPeriod
                    dtmStartDate = dtmPeriodEndDate
                End If
            Else
                'Period
                If xmlNode.nodeName = "INTERESTRATETYPE" Then
                    Set xmlPeriod = m_objXmlAssist.GetMandatoryNode(xmlNode, "INTERESTRATEPERIOD")
                Else
                    Set xmlPeriod = m_objXmlAssist.GetMandatoryNode(xmlNode, "PERIOD")
                End If

                intExpiredPeriod = CInt(xmlPeriod.Text)
                'MAR355 GHun
                If (intExpiredPeriod = -1) Then
                    blnFound = True
                    intRemainingPeriod = intTotalExpiredPeriod - intPeriodDifference
                    xmlPeriod.Text = CStr(intRemainingPeriod)
                Else
                'MAR355 End
                    intTotalExpiredPeriod = intTotalExpiredPeriod + intExpiredPeriod
                    If intTotalExpiredPeriod > intPeriodDifference Then
                        blnFound = True
                        intRemainingPeriod = intTotalExpiredPeriod - intPeriodDifference
                        xmlPeriod.Text = CStr(intRemainingPeriod)
                    Else
                        dtmStartDate = DateAdd("m", intExpiredPeriod, dtmStartDate)
                    End If
                End If
            End If
            
            If Not blnFound Then
                vxmlBandList.removeChild xmlNode
                Set xmlNode = vxmlBandList.firstChild
            End If
        Loop
    End If
      
    Set xmlNode = Nothing
    Set xmlPeriod = Nothing
    
    Exit Sub

ProcessPortedLoanBandsErr:

    Set xmlNode = Nothing
    Set xmlPeriod = Nothing

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
' PSC 18/06/2006 MAR1687 - End


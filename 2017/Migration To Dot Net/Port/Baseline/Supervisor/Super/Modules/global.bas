Attribute VB_Name = "Global"

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' Module        : Global
' Description   : Contains constants, enum's and global object variables
' Change history
' Prog      Date        Description
' DJP       09/06/01    Added SQLAssistSP for SQL Server port
' DJP       10/12/01    SYS2831 Client variants, added two Mortgage Product areas to ProductArea
'                       enum.
' STB       05/02/02    SYS3327 In-order to send UnitID to Batch, we must store it at logon.
' DJP       15/02/02    SYS4094 Remove reference to NoTab on LenderDetailsTabs enum.
' STB       07/03/02    SYS4208 Added grid From/To data verification function.
' STB       25/03/02    SYS4312 Added Optimus user's credentials as globals.
' CL        28/05/02    SYS4766 Merge MSMS & CORE
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' Module        : Global
' Description   : Contains constants, enum's and global object variables
' BMIDS Change history
' Prog      Date        Description
' GD        22/05/02    BMIDS0012
' AW        08/07/02    BMIDS00178  Removed the ODI envinronment global variables
' IK        07/03/03    BM0314 Added remove (DMS) Document Locks
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' MARS Change history
' GHun      16/11/2005  MAR312 Corrected OmigaUserTabs
' PSC       07/03/2006  MAR1298 Add POPULATE_KEYS_ORDERBY_PRIMARY to PopulateType
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' EPSOM Change history
' TW        09/10/2006  EP2_7 Added handling for Additional Borrowing Fee and Credit Limit Increase Fee
' PB        17/10/2006  EP2_14  E2CR20 - Unit organisation changes
' TW        17/10/2006  EP2_15 Introducer changes
' TW        30/11/2006  EP2_253 - Changes related to Mortgage Product Application Eligibility
' TW        11/12/2006  EP2_20 Added handling for Transfer of Equity Fee
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Option Explicit

' Global Objects
Public g_sVersion As String
Public g_Timer As Timer
Public g_clsMainSupport As New MainSuper
Public g_clsFormProcessing As New FormProcessing
Public g_clsDataAccess As New DataAccess
Public g_clsValidation As New Validation
Public g_clsGeneralAssist As New GeneralAssist
Public g_clsCombo As New ComboUtils
Public g_clsGlobalParameter As New GlobalParameter
Public g_clsErrorHandling As New ErrorHandling
Public g_clsSQLAssistSP As New SQLAssistSP
Public g_clsXMLAssist As New XMLAssistSP
Public g_sApplicationServer As String
Public g_sFunctionName As String
Public g_sVersionNumber As Long
Public Const SINGLE_LENDER = 1
Public g_clsVersion As New VersionControlTable
Public g_clsHandleUpdates As New HandleUpdates
Public g_bFirstTimeExecuted As Boolean
'PB 17/10/2006 EP2_14
'Used to indicate which specific option within the tree popup menu was chosen
Public g_intListViewDetail As LV_DETAIL
'PB 19/10/2006 EP2_13
'This global variable is used to avoid adding an optional parameter to a function in the
'TableAccess class, which would mean modifying every single class which implements it (lots).
Public g_lngResultCap As Long
'This one is used to prevent recordset validation
Public g_blnDoNotValidateRS As Boolean

'Global reference to the security manager component. This has to be kept for the
'duration of Supervisor as it will store a virtual table of securable items.
Public g_clsSecurityMgr As SecurityManager

' AW        08/07/02    BMIDS00178
'The ODI envinronment to connect to.
'Public g_sODIEnvironment As String
'Public g_sODIPassword As String
'Public g_bODISystemPresent As Boolean

Public Enum MSGReturnCode
    MSGSuccess = 1
    MSGFailure = 2
End Enum
Public Enum ControlOperation
    GET_CONTROL_VALUE = 0
    SET_CONTROL_VALUE = 1
End Enum

Public Enum SupervisorPromoteType
    PromoteEdit = 1
    PromoteDelete = 2
End Enum

'The currently logged in-user's credentials.
Public g_sSupervisorUser As String
Public g_sUnitID As String
Public g_sChannelID As String

Public Const OPT_YES = 0
Public Const OPT_NO = 1

Public Const REG_CONNECTIONS = SUPERVISOR_KEY + "\Connections"
Public Const REG_CONNECTION = REG_CONNECTIONS + "\Connection"
'Public Const REG_CONNECTION_ACTIVE = REG_CONNECTIONS + "\ActiveConnection"

Public Enum LenderDetailsTabs
    LenderDetails = 1
    ContactDetails
    AdditionalParameters
    LegalFees
    OtherFees
    MIGRateSets
    LenderPaymentDetails
End Enum

Public Enum ThirdPartyCombo
    ThirdPartyLender = 3
    ThirdPartyLegalRep = 10
    ThirdPartyValuer = 11
End Enum

Public Enum ThirdPartyType
    ThirdPartyInvalid = 0
    ThirdPartyLegalRepType
    ThirdPartyValuersType
    ThirdPartyLocalType
End Enum

Public Enum IntermediaryTabs
    IntDetails = 1
    IntContactDetails
    IntBankDetails
End Enum

Public Enum ThirdPartyDetailTabs
    TabThirdParty = 0
    TabContactDetails
    TabLegalRepDetails
    TabValuerDetails
    TabBankAccounts
End Enum

'AA Intermediary
Public Enum ProcurationFeeTabs
    TabNonProdSpecific = 1
    TabProdSpecific
    TabInsurance
    TabPackaging
End Enum

'MAR312 GHun
Public Enum OmigaUserTabs
    UserDetails = 1
    Qualifications = 2
    CompetencyHistory = 3
    UserRoles = 4
    SupervisorAccess = 5
End Enum
'MAR312 End

Public Enum ProductArea
    mortAreaProdDetails = 1
    mortAreaParameters
    mortAreaPurpOfLoan
    mortAreaChannelElig
    mortAreaEmpElig
    mortAreaSpecialGroup
    mortAreaTabAdminFees
    mortAreaTabValuationFees
    mortAreaOtherFee
    mortAreaTabIncentives
    mortAreaTabInterestRates
    mortAreaTabBaseRates
    mortAreaTypicalAPR
    mortAreaPropLocation
    mortAreaTypeOfBuyer
    mortAreaTypeOfApp
    mortAreaTabTTFees
    mortAreaTabIAFees
    mortAreaTabProdSwitchFees
    mortAreaTabAdditionalBorrowingFees  ' TW 09/10/2006 EP2_7
    mortAreaTabCreditLimitIncreaseFees  ' TW 09/10/2006 EP2_7
    mortAreaTabTransferOfEquityFees     ' TW 11/12/2006 EP2_20
' TW 30/11/2006 EP2_253
    mortAreaTabMortgageProductIncomeStatus
    mortAreaTabMortgageProductNatureOfLoan
    mortAreaTabMortgageProductProductClass
' TW 30/11/2006 EP2_253 End

End Enum

Public Enum ProductTabs
    mortProdDetails = 1
    mortProdParametersAndBands
    mortProdAppEligibility
    mortProdOtherEligibility
' TW 30/11/2006 EP2_253
    mortProdAppEligibility3
' TW 30/11/2006 EP2_253 End
    mortProdFees
    mortProdOtherFeesAndIncentives
' TW 09/10/2006 EP2_7
    mortProdFeeSet3
' TW 09/10/2006 EP2_7 End
    mortProdInteresRates
    mortProdMisc
    'GD BMIDS0012
    mortProdSpecialConditions
End Enum

Public Enum UnitTabs
    UnitDetails = 0
    UnitAddress
    UnitContact
    UnitTaskOwner
End Enum

' PSC 07/03/2006 MAR1298 - Start
Public Enum PopulateType
    POPULATE_FIRST_BAND = 1         ' For banded tables
    POPULATE_KEYS                   ' Populate based on the key values
    POPULATE_TABLE                  ' Just open the table - don't populate it
    POPULATE_ALL                    ' Populate everything in the table
    POPULATE_EMPTY                  ' Returns a recordset with no matches - useful for populating an empty datagrid (when adding, for example)
    POPULATE_KEYS_ORDERBY_PRIMARY   ' Populate based on the key values and order by the primary keys
End Enum
' PSC 07/03/2006 MAR1298 - End

'PB 23/10/2006 EP2_13 - Start
'Added enum to distinguish between two entity types which are represented by the same table,
'and ergo the same class.
'I have used the format OPTION_[table]_[entity] so that this can be reused in future cases.
Public Enum ClassOption
    OPTION_NONE = 0
    OPTION_PRINCIPALFIRM_PRINCIPALFIRM
    OPTION_PRINCIPALFIRM_PACKAGER
    OPTION_PRINCIPALFIRM_NETWORK
    OPTION_PRINCIPALFIRM_FIRMPACKAGER
    OPTION_PRINCIPALFIRM_FIRMBROKER
    OPTION_ARFIRM_INDIVIDUALPACKAGER
    OPTION_ARFIRM_INDIVIDUALBROKER
' TW 17/10/2006 EP2_15
    OPTION_CLUB_ASSOCIATION_CLUB
    OPTION_CLUB_ASSOCIATION_ASSOCIATION
' TW 17/10/2006 EP2_15 End
End Enum

Public Const DIRECTION_SOURCE = "S"
Public Const DIRECTION_DEST = "D"

Public Const MortgageProductCode As Integer = 1
Public Const MortgageProductStartDate As Integer = 2
Public Const MortgageProductVersion As Integer = 3

Public Const OBJECT_DESCRIPTION_KEY As String = "Supervisor Object Description"
Public Const OBJECT_EXTRA_VALUE As String = "Supervisor Extra Value"

Public Enum LockType
    Application = 1
    Customer
    ' ik_bm0314
    Document
End Enum

Public Enum ApplicationType
    OnLine = 1
    Offline = 2
End Enum

' DJP SQL Start
' Databases
Public Const DATABASE_SQL_SERVER = "SQL Server"
Public Const DATABASE_ORACLE = "Oracle"
Public Enum enumDatabaseTypes
    INDEX_ORACLE = 0
    INDEX_SQL_SERVER = 1
End Enum

Public Type ConnectionDetails
    sService As String
    sUserID As String
    sPassword As String
    sProvider As String
    sActive As String
    sAppServer As String
    sDatabaseServer As String
    sDatabaseType As String
End Type
' DJP SQL End

Public Enum DatabaseIndexes
    COL_DATABASE_TYPE = 1
    COL_SERVICE_NAME = 2
    COL_USER_ID = 3
    COL_PASSWORD = 4
    COL_APP_SERVER = 5
    COL_PROVIDER = 6
    COL_DATABASE_SERVER = 7
    COL_ACTIVE_CONNECTION = 8
End Enum

Public Type ComboGroup
    sValue As String
    sID As String
End Type
Public Type UpdateValues
    rs As ADODB.Recordset
    rowData As Variant
    clsTableAccess As TableAccess
    vBookMark As Variant
End Type

Public Type InterestRateTypes
    sValue As String
    sSaveValue As String
End Type
Public Enum IntermediaryType
    IntermediaryIndividual = 0
    IntermediaryOther
End Enum

Public Const constListViewLabel As Integer = 1

Private Declare Function GetComputerName Lib "kernel32" Alias "GetComputerNameA" (ByVal lpBuffer As String, nSize As Long) As Long


'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' Function      : GetMachineID
' Description   : Return the machine ID.
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''

' BMIDS Change history
' Prog      Date        Description
' GD        20/05/02    BMIDS : Added extra tab (Special Conditions) to frmProductDetails
Public Function GetMachineID() As String

    Dim lLength As Long
    Dim lReturn As Long
    Dim sMachineID As String
    
    sMachineID = Space$(255)
    lLength = Len(sMachineID)
    
    'Obtaine the computer name from the API.
    lReturn = GetComputerName(sMachineID, lLength)
    sMachineID = Left$(sMachineID, lLength)
    
    'Return the MachineID.
    GetMachineID = sMachineID

End Function

Public Sub SetSpashText(sText As String)
        
    frmSplash.Show vbModeless
    frmSplash.lblStatusTitle = sText
    frmSplash.lblVersion = g_sVersion
    frmSplash.Refresh
End Sub

Public Sub HideSplash()
    Unload frmSplash
End Sub

Public Sub ValidateRecordset(rs As ADODB.Recordset, sDescription As String)
    If rs Is Nothing Then
        g_clsErrorHandling.RaiseError errRecordSetEmpty, sDescription
    End If
End Sub

Public Function GetApplicationServer() As String
    On Error GoTo Failed
    
    Dim clsConn As SupervisorConnection
    g_clsDataAccess.GetSupervisorConnection clsConn
    
    GetApplicationServer = clsConn.GetAppServer()
    
    Exit Function
Failed:
    g_clsErrorHandling.RaiseError Err.Number, Err.DESCRIPTION
End Function

Public Function GetObjectContext() As Object
    Set GetObjectContext = Nothing
End Function

Public Function TableAccess(clsTableAccess As TableAccess) As TableAccess
    Set TableAccess = clsTableAccess
End Function

Public Function BandedTable(clsBandedTable As BandedTable) As BandedTable
    Set BandedTable = clsBandedTable
End Function

Public Function ValidateDateColumns(ByRef grdGrid As MSGDataGrid, ByVal iFromColumnIdx As Integer, ByVal iToColumnIdx As Integer) As Boolean

    Dim sToDate As String
    Dim sFromDate As String
    Dim iCurrentRow As Integer
    
    'Iterate through all the rows of the grid control.
    For iCurrentRow = 0 To (grdGrid.Rows - 1)
        'Extract the to and from dates.
        sToDate = grdGrid.GetAtRowCol(iCurrentRow, iToColumnIdx - 1)
        sFromDate = grdGrid.GetAtRowCol(iCurrentRow, iFromColumnIdx - 1)
    
        'Verify the dates.
        If ValidateDate(sFromDate, sToDate) = False Then
            'Place the cursor in the invalid From date.
            grdGrid.col = (iFromColumnIdx - 1)
            grdGrid.Row = iCurrentRow
            grdGrid.SetGridFocus
            
            'Raise a warning and stop further processing.
            g_clsErrorHandling.RaiseError errGeneralError, "The From date cannot be after the To date.", ErrorUser
        End If
    Next iCurrentRow
    
    ValidateDateColumns = True

End Function

Public Function ValidateDate(ByVal sFromDate As String, ByVal sToDate As String) As Boolean

    'If either of the dates is missing then don't both trying to validate them.
    If (sFromDate = "") Or (sToDate = "") Then
        ValidateDate = True
    
    'VB doesn't support progressive expression checking...
    ElseIf CDate(sFromDate) <= CDate(sToDate) Then
        ValidateDate = True
    End If

End Function


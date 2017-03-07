VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "MainSuper"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' Class         : MainSuper
' Description   : Support class for the main dialog, frmMain.
' History:
'
' Prog      Date        Description
' DJP       09/11/00    Phase 2 Task Management
' CL        24/04/01    Modification to allow for inclusion of frmEditCurrencies
' DJP       28/06/01    SQL Server port
' AA        28/06/01    Intermediaries
' DJP       03/09/01    SYS2564 Clear table history before reinitialising Supervisor
' DJP       04/08/01    Added logging to aid debugging of errors
' DJP       06/08/01    Added code to disable currencies if tables don't exist.
' DJP       22/10/01    SYS2831 Updated for Client variants.
' STB       15/11/01    SYS2550: Made HandleReturn public to be used by Intermediaries.
' DJP       22/11/01    SYS2912 SQL Server locking problem. Added CreateLender and removed
'                       references to frmMainLenderDetails
' DJP       03/12/01    SYS2912 SQL Server locking problem. Added CreateProduct and removed
'                       references to frmMainProductDetails
' STB       04/12/01    SYS2912 SQL Server locking problem. Removed references to frmMainUser
'                       and frmMainThirdParty. Also corrected a bug with ThirdParty enum.
' STB       07/12/01    SYS1942 - HandleReturn() will now refresh the listview regardless
'                       of a form's return code.
' STB       21/12/01    SYS2550 - Integrate Intermediaries.
' DJP       09/02/02    SYS2831 Support client variants for Mortgage Product
' SA        18/01/02    SYS3327 - HandleBatchProcessing changed.
' STB       21/01/02    SYS2957 Supervisor Security Enhancement.
' DJP       13/02/02    SYS4052 Allow direct access to Supervisor if no users exist.
' DJP       14/02/02    SYS4093 If the user hasn't logged on, reset the UserID.
' DJP       24/02/02    SYS4142 - TreeView inheritence, ListView inheritence.
' DJP       24/02/02    SYS4148 - Deletion inheritence.
' STB       05/04/02    SYS3149 - Deleting a task should also delete stage tasks.
' STB       05/04/02    SYS4247 - Deleting all the Baserate Bands for a Set should remove remove the set.
' STB       13/05/02    SYS4417 - Added AllowableIncomeFactors.
' CL        22/05/02    SYS4707 - Modification to DeleteObject.
' BG        10/06/02    SYS4692 - Added code to DeleteObject for valuation fees and admin fees to call
'                                   new method DeleteIfNoBands
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''

'   BMIDS
'   AW      14/05/02    BM088   Added Income Multiple processing
'   AW      16/05/02    BM087   Added Mig rate set processing
'   AW      21/05/02    BM017   Added redemption fee set processing
'   AW      18/05/02    BM00023 - Delete the parent set record if all bands have been deleted.
'   AW      28/05/02    BM00016 - When promoting deletes, need to pass the target database into
'                                  DeleteNoBands
'   MO      03/07/2002  BMIDS00054  Added new items to the Products popup menu and modified function HandleProducts()
'   AW      11/07/02    BMIDS00177  Removed LV_LENDERS_LEDGER_CODES from HandleLenders()
'   CL      01/10/02    BMIDS00201 Added LV_PRODUCTS_SPECIAL_CONDITIONS
'   MV      15/01/02    BM0085  Added a New Form frmAddComboGroup  Amended HandleTreeNew()
'   DJP     20/02/03    BM0318 Functionality for Third Party searching
'   DJP     24/02/03    BM0318 Functionality for Third Party searching
'   DJP     24/02/03    BM0318 Functionality for Third Party searching
'   DJP     05/03/03    BM0282 Searches for Users, Products, Locks and Conditions
'   IK      07/03/03    BM0314 Added remove (DMS) Document Locks
'   BS     25/03/03     BM0282 Stop text repeating in listview title bar
'   GHun    02/04/2003  BM0425 Commented out GetBatchSchedule as it is no longer used
'   BS      28/04/03    BM0240 Change tests for Conditional Compilation Arguments for automated build
'   MC      02/06/2004  BMIDS763    BMIDS763 - CC075 - FEESETS Added (Product Switch fee/Insurance Admin fEE/TT fee sets)
'   JD      15/06/04    BMIDS765    CC076 addition of Rental Income Rate Sets
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' MARS History
' Prog      Date        Description
' GHun      16/08/2005  MAR45 Apply BBG1370 (New screen for print configuration)
' PJO       28/11/2005  MAR81 Solicitor Panel Maintenance
' TK        30/11/2005  MAR81 HandleActivateThirtParty
' GHun      14/10/2005  MAR202 New screen for printing packs
' GHun      04/04/2005  MAR1300 Delete linked tasks when deleting a task
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' EPSOM History
' PB        24/04/2006  EP367   Changes merged into Epsom Supervisor
' PB        10/05/2006  EP522   "Can't edit a newly added Packager"
' PB        15/05/2006  EP375   Allow all Supervisor administrators to set the legal rep status
' TW        09/10/2006  EP2_7   Added handling for Additional Borrowing Fees and Credit Limit Increase Fees
' PB        17/10/2006  EP2_14  E2CR20 - Unit organisation changes
' TW        17/10/2006  EP2_15  Club/Association processing
' PB        01/11/2006  EP2_13  Changes for 'Find broker/Packager' functions
' TW        18/11/2006  EP2_132 ECR20/21 Changes for promotion of deletes
' TW        23/11/2006  EP2_172 Change control EP2_5 - E2CR16 changes related to Introducer/Product Exclusives
' TW        30/11/2006  EP2_253 Changes related to Mortgage Product Application Eligibility
' TW        06/12/2006  EP2_330 Promotions
' TW        11/12/2006  EP2_20  WP1 - Loans & Products Supervisor Changes part 3
' TW        14/12/2006  EP2_518 Procuration Fees Support
' TW        18/12/2006  EP2_568 Add functionality to select which Omiga Users are returned
' TW        15/01/2007  EP2_826 Rationalisation of pop-up menus and actions to improve consistency and usability
' TW        16/01/2007  EP2_859 - Principal Firms/Network display and search
' TW        30/01/2007  EP2_1122 - Error on Third Parties - Other
' TW        01/02/2007  EP2_1036 - Principal Firms/Network display and search (Follow-on)
' TW        03/02/2007  EP2_1101 - Promotion inconsistencies
' TW        05/02/2007  EP2_706 - Should  network be mandatory data for ar firms.
' TW        22/02/2007  EP2_1577 - Mortgage Product "Errors" pop-up not displayed
' TW        23/03/2007  EP2_1942 - Promotion of Income Factors does not work
' TW        21/04/2007  EP2_2497 - Find Users - The find facility returns an error "Invalid Property Value" when searching for user
' TW        27/11/2007  DBM594 - Add new functionality "Payments for Completion"
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Option Explicit

' Private data
Private m_rsMain As ADODB.Recordset
Private m_bUpdated As Boolean
Private m_nCurrentIndex As Long
Private m_clsIntermediary As Intermediary
Private m_bLogging As Boolean

' Constants
Private Const PRODUCTS_EDIT As Integer = 1
Private Const PRODUCTS_MAINT As Integer = 2
Private Const ARG_LOGGING = "LOG"
Private Const LOG_PATH = "Supervisor.log"

Private Type SUPERVISOR_OBJECTS
    bTaskManagement As Boolean
    bBatch As Boolean
    bPrintingTemplate As Boolean
    bIntermediary As Boolean
    bBaseRate As Boolean
    bBusinessGroups As Boolean
    bCurrencies As Boolean
End Type

Private supervisorObjects As SUPERVISOR_OBJECTS

' Enum's and types
Public Enum LIST_STATE
    SET_SELECTION = 1
    GET_SELECTION
    CLEAR_SELECTION
End Enum
Friend Sub SetRecordSet(rs As ADODB.Recordset)
    Set m_rsMain = rs
End Sub

''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' Function      : HandleTreeFind
' Description   : Called when the user selects to find an item from the treeview.
'                 Loads the dialog appropriate for the item selected by the user.
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Friend Sub HandleTreeFind()
    On Error GoTo Failed
    Dim vKey As Variant
    Dim frmFind As Form
    Dim sHeaderTitle As String
    Dim clsTableAccess As TableAccess
    Dim bModal As Boolean
    
    bModal = True
    ' Which item do we want to search for?
    vKey = frmMain.GetSelectedTreeKey()
    
' TW 21/04/2007 EP2_2497
' Ensure headers are set up otherwise "Invalid Property Value" is returned
    frmMain.LoadListViewHeaders
' TW 21/04/2007 EP2_2497 End
    
    If (Not IsNull(vKey)) Then
                    
        Select Case vKey
        Case LENDERS
        Case ADMIN_FEES
        Case VALUATION_FEES
        Case BASE_RATES
        Case INCOME_MULTIPLE
        Case MP_MIG_RATE_SETS
        Case REDEM_FEE_SETS
        Case LENDER_ADDRESS
        Case GLOBAL_PARAM_FIXED
        Case GLOBAL_PARAM_BANDED
        Case ERROR_MESSAGES
        Case COMBOBOX_ENTRIES
        Case DISTRIBUTION_CHANNELS
        Case COUNTRIES
        Case DEPARTMENTS
            Set clsTableAccess = New DepartmentTable
            Set frmFind = frmFindDepartments
            sHeaderTitle = DEPARTMENTS
        
        Case UNITS
            Set clsTableAccess = New UnitTable
            Set frmFind = frmFindUnits
            sHeaderTitle = UNITS
        Case REGIONS
        
        Case MORTGAGE_PRODUCTS
            Set clsTableAccess = New MortgageProductTable
            Set frmFind = frmFindProducts
            sHeaderTitle = MORTGAGE_PRODUCTS
        
        Case APPLICATION_PROCESSING
            Set clsTableAccess = New ApplicationTable
            Set frmFind = frmFindApplication
            
            sHeaderTitle = APPLICATION_PROCESSING
            bModal = False
        
        ' DJP BM0282
        Case CONDITIONS
            Set clsTableAccess = New ConditionsTable

            Set frmFind = frmFindConditions

            sHeaderTitle = CONDITIONS
        
        ' DJP BM0318 -- Start -- Need a find screen for Third Parties
        Case LOCAL_ADDRESS
            Set clsTableAccess = New PanelTable

            Set frmFind = frmFindPanel
            frmFindPanel.SetThirdPartyType ThirdPartyLocalType

            sHeaderTitle = LOCAL_ADDRESS

        Case VALUER_ADDRESS
            Set clsTableAccess = New PanelTable

            Set frmFind = frmFindPanel
            frmFindPanel.SetThirdPartyType ThirdPartyValuersType

            sHeaderTitle = PANEL_ADDRESS
        
        Case LEGAL_REP_ADDRESS
            Set clsTableAccess = New PanelTable

            Set frmFind = frmFindPanel
            frmFindPanel.SetThirdPartyType ThirdPartyLegalRepType
            
            sHeaderTitle = PANEL_ADDRESS
        
        ' DJP BM0282
        Case USERS
            Set clsTableAccess = New OmigaUserTable

            Set frmFind = frmFindUsers
            
            sHeaderTitle = USERS
        
        'PB EP2_13 07/11/2006 Begin
        'New treeview menu structure
        Case ASSOCIATIONS
            
            Set clsTableAccess = New AssociationTable
            
            Set frmFind = frmFindAssociation
            
            sHeaderTitle = ASSOCIATIONS
            
        Case CLUBS
            
            Set clsTableAccess = New ClubTable
            
            Set frmFind = frmFindClubs
            
            sHeaderTitle = CLUBS
        Case PRINCIPALS
            Set clsTableAccess = New PrincipalFirmTable
            Set frmFind = New frmFindPrincipalFirm
        
        Case PACKAGERS

            Set clsTableAccess = New FirmPackagerTable

            Set frmFind = frmFindFirmPackager
            
            sHeaderTitle = PACKAGER & " " & FIRMS
        
        Case ARFIRMS
            Set clsTableAccess = New ARFirmTable
            Set frmFind = New frmFindARFirm
        
            
        Case INDIVIDUALS & "_" & PACKAGERS
            Set clsTableAccess = New IndividualPackagerTable
            Set frmFind = New frmFindIndividualPackager
        
        Case ARBROKERS
            Set clsTableAccess = New ARBrokerTable
            Set frmFind = New frmFindARBroker
            
        Case DABROKERS
            Set clsTableAccess = New DABrokerTable
            Set frmFind = New frmFindDABroker
            
        'PB EP2_13 End
        
        ' DJP -- End -- BM0318 Start Need a find screen for Third Parties
        
' TW 17/10/2006 EP2_15
    ' PB 01/11/2006 EP2_13
'        Case FIRMS & "_" & PACKAGERS
'            Set clsTableAccess = New IndividualPackagerTable
'            Set frmFind = New frmFindIndividualPackager
            
'        Case FIRMS & "_" & BROKER
'
'            Set clsTableAccess = New FirmBrokerTable
'
'            Set frmFind = frmFindFirmBroker
'
'            sHeaderTitle = BROKER & " " & FIRMS
        
            
'        Case ARFIRMS
'
'            Set clsTableAccess = New ARFirmTable
'
'            Set frmFind = frmFindARFirm
            
            
    ' PB 01/11/2006 EP2_13 End
' TW 17/10/2006 EP2_15 End
        
        
        Case Else
            Dim clsTreeNewCS As TreeViewCS
            Set clsTreeNewCS = New TreeViewCS
            
            clsTreeNewCS.HandleTreeFind vKey, frmFind, clsTableAccess

        End Select
    
        If Not frmFind Is Nothing Then
            frmFind.SetTableClass clsTableAccess
            BeginWaitCursor
            
            If bModal Then
                frmFind.Show vbModal, frmMain
                
                If Not clsTableAccess Is Nothing Then
                    If clsTableAccess.GetUpdated() = True Then
                        frmMain.SetColumnHeader sHeaderTitle
                        g_clsFormProcessing.PopulateFromRecordset frmMain.lvListView, clsTableAccess
                    End If
                End If
                Unload frmFind
            Else
                'frmFind.Show vbModeless, frmMain
                frmFind.Show vbModeless, frmMain
                
                SetTopMostWindow frmFind.hWnd, True
            End If
            
        Else
            g_clsErrorHandling.RaiseError errGeneralError, "Unable to show Find dialog - Form is empty"
        End If
    End If
    Exit Sub
Failed:
    g_clsErrorHandling.RaiseError Err.Number, Err.DESCRIPTION
End Sub
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' Function      : HandleTreeNew
' Description   : Called when the user right clicks on the treeview and selects
'                 "New".
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Friend Sub HandleTreeNew(vKey As Variant)
    On Error GoTo Failed
    
    Dim frmNew As Form
    Dim enumReturn As MSGReturnCode
    Dim clsTreeNewCS As TreeViewCS
    
    CloseFindDialogs
    enumReturn = MSGFailure
    
    If (Not IsNull(vKey)) Then
                    
        Select Case vKey
        Case LENDERS
            enumReturn = CreateLender
        
        Case MORTGAGE_PRODUCTS
            enumReturn = CreateProduct
        
' TW 09/10/2006 EP2_7
        Case ADDITIONAL_BORROWING_FEES
            Set frmNew = frmEditAdditionalBorrowingFees
            
        Case CREDIT_LIMIT_INCREASE_FEES
            Set frmNew = frmEditCreditLimitIncreaseFees
' TW 09/10/2006 EP2_7 End
' TW 11/12/2006 EP2_20
        Case TRANSFER_OF_EQUITY_FEES
            Set frmNew = frmEditTransferOfEquityFees
' TW 11/12/2006 EP2_20 End
' TW 14/12/2006 EP2_518
        Case DEFAULT_PROCURATION_FEES
            Set frmNew = frmEditDefaultProcurationFees
        Case LOAN_AMOUNT_ADJUSTMENTS
            Set frmNew = frmEditLoanAmountAdjustments
        Case LTV_AMOUNT_ADJUSTMENTS
            Set frmNew = frmEditLTVAmountAdjustments
' TW 14/12/2006 EP2_518 End
        
        Case ADMIN_FEES
            Set frmNew = frmEditAdminFees
        
        '*=[MC]BMIDS763 - CC075 - FEESETS
        Case PRODUCT_SWITCH_FEESETS
            Set frmNew = frmEditProductSwitchFee
        Case INSURANCE_ADMIN_FEESETS
            Set frmNew = frmEditInsuranceAdminFee
        Case TT_FEESETS
            Set frmNew = frmEditTTFee
        '*=[MC]BMIDS763 - SECTION END
        
        'JD BMIDS765
        Case RENTAL_INCOME_RATES
            Set frmNew = frmEditRentalIncomeRateSet
        Case VALUATION_FEES
            Set frmNew = frmEditValuation
        
        Case BASE_RATES
            Set frmNew = frmEditBaseRates
        '   AW  14/05/02    BM088
        Case INCOME_MULTIPLE
            Set frmNew = frmEditIncMultiples
                 
        '   AW  16/05/02    BM087
        Case MP_MIG_RATE_SETS
            Set frmNew = frmEditMPMigRates
            
        '   AW  16/05/02    BM017
        Case REDEM_FEE_SETS
            Set frmNew = frmEditRedemFeeSets
                    
        Case PAYMENT_PROTECTION_RATES
            Set frmNew = frmEditPayProtRates
        
        Case LOCAL_ADDRESS
            frmEditThirdParty.SetThirdPartyType ThirdPartyLocalType
            Set frmNew = frmEditThirdParty
        
        Case LEGAL_REP_ADDRESS
            frmEditThirdParty.SetThirdPartyType ThirdPartyLegalRepType
            Set frmNew = frmEditThirdParty
        
        Case VALUER_ADDRESS
            frmEditThirdParty.SetThirdPartyType ThirdPartyValuersType
            Set frmNew = frmEditThirdParty
        
        Case GLOBAL_PARAM_FIXED
            Set frmNew = frmEditGlobalFixed
        
        Case GLOBAL_PARAM_BANDED
            Set frmNew = frmEditGlobalBanded

        Case ERROR_MESSAGES
            Set frmNew = frmErrors
        
        Case COUNTRIES
            Set frmNew = frmEditCountry

        Case COMBOBOX_ENTRIES
            Set frmNew = frmAddComboGroup
        
        Case DISTRIBUTION_CHANNELS
            Set frmNew = frmEditDistChannel
    
        Case DEPARTMENTS
            Set frmNew = frmEditDepartment
        
        Case REGIONS
            Set frmNew = frmEditRegion
        
        Case UNITS
            Set frmNew = frmEditUnit
        
        Case COMPETENCIES
            Set frmNew = frmEditCompetency
        
        Case WORKING_HOURS
            Set frmNew = frmEditWorkingHoursType
        
        Case USERS
            Set frmNew = frmEditUser

        Case LIFE_COVER_RATES
            Set frmNew = frmEditLifeCover

        Case BUILDINGS_AND_CONTENTS_PRODUCTS
            Set frmNew = frmEditBuildingsAndContents

        Case PAYMENT_PROTECTION_PRODUCTS
            Set frmNew = frmEditPayProtProducts
        
        ' --- Phase 2 Task Management Start
        Case TASK_MANAGEMENT_TASKS
            Set frmNew = frmEditTasks
        Case TASK_MANAGEMENT_ACTIVITIES
            Set frmNew = frmEditActivities
        Case TASK_MANAGEMENT_STAGES
            Set frmNew = frmEditStages
        ' --- Phase 2 Task Management End
        Case BUSINESS_GROUPS
            Set frmNew = frmEditBusinessGroup
        Case ADDITIONAL_QUESTIONS
            Set frmNew = frmEditQuestions
            
        Case CONDITIONS
            Set frmNew = frmEditConditions
        Case BATCH_SCHEDULER
            Set frmNew = frmEditBatch
        Case BASE_RATE
            Set frmNew = frmEditRate
        Case PRINTING_TEMPLATE
            Set frmNew = frmEditPrintingTemplate
        'MAR45 GHun
        Case PRINTING_DOCUMENT
            Set frmNew = frmEditPrintingDocument
        'MAR45 End
        'MAR202 GHun
        Case PRINTING_PACK
            Set frmNew = frmEditPrintingPack
        'MAR202 End
        Case CURRENCIES
            Set frmNew = frmEditCurrencies
        Case INCOME_FACTORS
            Set frmNew = frmEditIncomeFactors
        Case INDIVIDUAL, INTERMEDIARIES_COMPANY, ADMIN_CENTRE, LEADAGENT
        
' TW 17/10/2006 EP2_15
        Case ASSOCIATIONS, CLUBS
            Set frmNew = frmEditAssociation
        
        Case PRINCIPALS, PACKAGERS
            Set frmNew = frmEditPackager
            
        Case ARFIRMS
            Set frmNew = frmEditBroker
        
        Case INDIVIDUALS & "_" & PACKAGERS
            Set frmNew = frmEditIndividualPackager
        
        Case ARBROKERS
            Set frmNew = frmEditIndividualARBroker
        
        Case DABROKERS
            Set frmNew = frmEditIndividualDABroker
' TW 17/10/2006 EP2_15 End
        
        Case Else
            ' Check for any client specific New items.
            Set clsTreeNewCS = New TreeViewCS
            clsTreeNewCS.HandleTreeNew vKey, frmNew, enumReturn
        
        End Select
            
        If Not frmNew Is Nothing Then
            frmNew.SetIsEdit False
            
            g_clsDataAccess.BeginTrans
            
            frmNew.Show vbModal, frmMain
                        
            enumReturn = frmNew.GetReturnCode()
            
            Unload frmNew
        End If

        If enumReturn = MSGSuccess Then
            Set frmMain.tvwDB.SelectedItem = frmMain.tvwDB.Nodes(vKey)
            frmMain.SetTreeStatus
        End If
        
        HandleReturn enumReturn

    End If
    Exit Sub
Failed:
    g_clsDataAccess.RollbackTrans
    g_clsDataAccess.CloseConnection
    g_clsErrorHandling.RaiseError Err.Number, Err.DESCRIPTION
    
    Exit Sub
    Resume
End Sub
Public Sub SelectItem(lvListView As MSGListView, Optional nItem As LIST_STATE = GET_SELECTION)
    On Error GoTo Failed
    
    SetSelectedItem lvListView, m_nCurrentIndex, nItem
    
    Exit Sub
Failed:
    g_clsErrorHandling.RaiseError Err.Number, Err.DESCRIPTION
End Sub
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' Function      : SetSelectedItem
' Description   : Sets the selection of the listview to the current index.
'                 This is so that when the user has edited an item, we
'                 can a) repopulate to reflect the changes, and b) set the
'                 selection to what it was before the edit
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Public Sub SetSelectedItem(lvListView As MSGListView, nCurrentIndex As Long, Optional nItem As LIST_STATE = GET_SELECTION)
    On Error GoTo Failed
    Dim Item As ListItem
    Dim nCount As Long
    
    nCount = lvListView.ListItems.Count
    
    If nItem = GET_SELECTION Then
        Set Item = lvListView.SelectedItem
        
        If Not Item Is Nothing Then
            nCurrentIndex = Item.Index
        Else
            nCurrentIndex = -1
        End If
    ElseIf nItem = SET_SELECTION Then
        If nCount > 0 Then
            If nCurrentIndex <= 0 Then
                nCurrentIndex = 1
                ' Could be invalid if the item has been removed
            ElseIf nCurrentIndex > nCount Then
                nCurrentIndex = nCount
            End If
            
            ' Clear all first
            Dim lstItem As ListItem
            For Each lstItem In lvListView.ListItems
                lstItem.Selected = False
            Next
            
            Set Item = lvListView.ListItems.Item(nCurrentIndex)
            Set lvListView.SelectedItem = Item
            
            Item.EnsureVisible
            
' TW 15/01/2007 EP2_826
'            frmMain.EnableEdit
' TW 15/01/2007 EP2_826 End
        End If
    Else
        ' Clear
        Set lvListView.SelectedItem = Nothing
        
        For Each lstItem In lvListView.ListItems
            lstItem.Selected = False
        Next
    
    End If
    Exit Sub
Failed:
    g_clsErrorHandling.RaiseError Err.Number, Err.DESCRIPTION
End Sub
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' Function      : HandleUsers
' Description   : The user screen has a number of tabs. The purpose of this function
'                 is to load the user screen set at the tab the user selected by right
'                 clicking on the listview when users are displayed, and selecting the
'                 tab they want (specified in enumTab)
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Public Function HandleUsers(nIndex As Integer, Optional enumTab As OmigaUserTabs) As MSGReturnCode
    On Error GoTo Failed
    Dim colMatchValues As New Collection
    Dim enumReturn As MSGReturnCode
    
    enumReturn = MSGFailure
    
    SetSelectedItem frmMain.lvListView, m_nCurrentIndex
    
    Select Case nIndex
    Case LV_USER_EDIT
        GetListViewKeys frmMain.lvListView, colMatchValues
        frmEditUser.SetKeys colMatchValues
        frmEditUser.SetIsEdit True
        
        frmEditUser.SetActiveTab enumTab
        
        On Error GoTo EditFailed
        g_clsDataAccess.BeginTrans
        
        frmEditUser.Show vbModal, frmMain
        enumReturn = frmEditUser.GetReturnCode()
        Unload frmEditUser
        HandleReturn enumReturn

    Case LV_USER_DELETE
        DeleteRecord
    
    End Select
    
    SetSelectedItem frmMain.lvListView, m_nCurrentIndex, SET_SELECTION
    HandleUsers = enumReturn
    
    Exit Function
Failed:
    g_clsDataAccess.CloseConnection
    g_clsErrorHandling.RaiseError Err.Number, Err.DESCRIPTION
    Exit Function

EditFailed:
    g_clsDataAccess.RollbackTrans
    g_clsDataAccess.CloseConnection
    g_clsErrorHandling.RaiseError Err.Number, Err.DESCRIPTION
End Function
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' Function      : HandleEditThirtParty
' Description   : There are a number of third party screens which are all very
'                 similar. This function loads the required third party screen
'                 as specified when the user right clicked on the listview when
'                 displaying a list of the third party in question (specified by
'                 the enum Tab parameter)
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Public Function HandleEditThirtParty(Optional enumTab As ThirdPartyDetailTabs) As MSGReturnCode
    On Error GoTo Failed
    Dim vKey As Variant
    Dim colMatchValues As New Collection
    Dim enumThirdPartyType As ThirdPartyType
    Dim enumReturn As MSGReturnCode
    
    enumReturn = MSGFailure
    SetSelectedItem frmMain.lvListView, m_nCurrentIndex
    vKey = frmMain.GetSelectedTreeKey()

    If (Not IsNull(vKey)) Then
        Select Case vKey
            Case LOCAL_ADDRESS
                enumThirdPartyType = ThirdPartyLocalType
            
            Case LEGAL_REP_ADDRESS
                enumThirdPartyType = ThirdPartyLegalRepType
            
            Case VALUER_ADDRESS
                enumThirdPartyType = ThirdPartyValuersType
            
            Case Else
                g_clsErrorHandling.RaiseError errGeneralError, "EditAddress: Unknown Address " & vKey
        End Select
            
        GetListViewKeys frmMain.lvListView, colMatchValues
        
        frmEditThirdParty.SetKeys colMatchValues
        frmEditThirdParty.SetIsEdit True
        frmEditThirdParty.SetThirdPartyType enumThirdPartyType
        
        If Not IsMissing(enumTab) Then
            frmEditThirdParty.SetActiveTab enumTab
        End If
        g_clsDataAccess.BeginTrans
        frmEditThirdParty.Show vbModal, frmMain

        enumReturn = frmEditThirdParty.GetReturnCode()
        Unload frmEditThirdParty
    
        HandleReturn enumReturn
    End If
    
    HandleEditThirtParty = enumReturn
    SetSelectedItem frmMain.lvListView, m_nCurrentIndex, SET_SELECTION
    Exit Function
Failed:
    g_clsDataAccess.RollbackTrans
    g_clsDataAccess.CloseConnection
    g_clsErrorHandling.RaiseError Err.Number, Err.DESCRIPTION
End Function
'TK 29/11/2005 MAR81
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' Function      : HandleActivateThirtParty
' Description   : There are a number of third party screens which are all very
'                 similar. This function loads the required third party screen
'                 as specified when the user right clicked on the listview when
'                 displaying a list of the third party in question (specified by
'                 the enum Tab parameter)
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Public Function HandleActivateThirtParty(Optional enumTab As ThirdPartyDetailTabs) As MSGReturnCode
    On Error GoTo Failed
    Dim vKey As Variant
    Dim colMatchValues As New Collection
    Dim enumThirdPartyType As ThirdPartyType
    Dim enumReturn As MSGReturnCode
    Dim sStatus As String
    Dim sLegalRepUserID As String
    
    ' PB 15/05/2006 EP375 - commented out next 4 lines
    'If frmMain.lvListView.ListItems.Item(frmMain.lvListView.SelectedItem.Index).SubItems(6) = g_sSupervisorUser Then
    '    MsgBox "You are not able to Active this Legal Rep as you were the last person to update it."
    '    Exit Function
    'End If

    enumReturn = MSGFailure
    SetSelectedItem frmMain.lvListView, m_nCurrentIndex
    sStatus = frmMain.lvListView.getValueFromName("Legal Rep Status")
    'sLegalRepUserID = frmMain.lvListView.getValueFromName("Legal Rep UserID")
    vKey = frmMain.GetSelectedTreeKey()
    If (Not IsNull(vKey)) Then
        Select Case vKey
            Case LOCAL_ADDRESS
                enumThirdPartyType = ThirdPartyLocalType
            
            Case LEGAL_REP_ADDRESS
                enumThirdPartyType = ThirdPartyLegalRepType
            
            Case VALUER_ADDRESS
                enumThirdPartyType = ThirdPartyValuersType
            
            Case Else
                g_clsErrorHandling.RaiseError errGeneralError, "EditAddress: Unknown Address " & vKey
        End Select
            
        GetListViewKeys frmMain.lvListView, colMatchValues
        
        frmEditThirdParty.SetKeys colMatchValues
        frmEditThirdParty.SetIsEdit True
        frmEditThirdParty.SetIsActivate True
        
        frmEditThirdParty.SetLegalRepStatus sStatus
        frmEditThirdParty.SetLegalRepUserID sLegalRepUserID
    
        frmEditThirdParty.SetThirdPartyType enumThirdPartyType
        
        If Not IsMissing(enumTab) Then
            frmEditThirdParty.SetActiveTab enumTab
        End If
        g_clsDataAccess.BeginTrans
        frmEditThirdParty.Show vbModal, frmMain

        enumReturn = frmEditThirdParty.GetReturnCode()
        Unload frmEditThirdParty
    
        HandleReturn enumReturn
    End If
    
    HandleActivateThirtParty = enumReturn
    SetSelectedItem frmMain.lvListView, m_nCurrentIndex, SET_SELECTION
    Exit Function
Failed:
    g_clsDataAccess.RollbackTrans
    g_clsDataAccess.CloseConnection
    g_clsErrorHandling.RaiseError Err.Number, Err.DESCRIPTION
End Function

''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' Function      : ListEdit
' Description   : Called when the user selects edit (from the right click menu)
'                 on the listview. Need to find the keys assocaited with the
'                 row clicked, and load the appropriate form with the correct
'                 table information
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Public Sub ListEdit()
    On Error GoTo Failed
    Dim vKey As Variant
    Dim colMatchValues As Collection
    Dim frmEdit As Form
    Dim enumReturn As MSGReturnCode
    Dim bCloseFind As Boolean
    
    enumReturn = MSGFailure
    bCloseFind = True
    
    vKey = frmMain.GetSelectedTreeKey()
    
    BeginWaitCursor
    
    If (Not IsNull(vKey)) Then
        If Not frmMain.lvListView.SelectedItem Is Nothing Then
            SetSelectedItem frmMain.lvListView, m_nCurrentIndex
            GetListViewKeys frmMain.lvListView, colMatchValues
            
            Select Case vKey

' TW 09/10/2006 EP2_7
            Case ADDITIONAL_BORROWING_FEES
                Set frmEdit = frmEditAdditionalBorrowingFees
                
            Case CREDIT_LIMIT_INCREASE_FEES
                Set frmEdit = frmEditCreditLimitIncreaseFees
' TW 09/10/2006 EP2_7 End
' TW 11/12/2006 EP2_20
            Case TRANSFER_OF_EQUITY_FEES
                Set frmEdit = frmEditTransferOfEquityFees
' TW 11/12/2006 EP2_20 End
' TW 14/12/2006 EP2_518
            Case DEFAULT_PROCURATION_FEES
                Set frmEdit = frmEditDefaultProcurationFees
            Case LOAN_AMOUNT_ADJUSTMENTS
                Set frmEdit = frmEditLoanAmountAdjustments
            Case LTV_AMOUNT_ADJUSTMENTS
                Set frmEdit = frmEditLTVAmountAdjustments
' TW 14/12/2006 EP2_518 End
            
            Case ADMIN_FEES
                Set frmEdit = frmEditAdminFees
            
            '*=[MC]BMIDS763 - CC075 - FEESETS
            Case PRODUCT_SWITCH_FEESETS
                Set frmEdit = New frmEditProductSwitchFee
            Case INSURANCE_ADMIN_FEESETS
                Set frmEdit = frmEditInsuranceAdminFee
            Case TT_FEESETS
                Set frmEdit = frmEditTTFee
            '*=[MC]SECTION END - BMIDS763
            'JD BMIDS765
            Case RENTAL_INCOME_RATES
                Set frmEdit = frmEditRentalIncomeRateSet
            Case VALUATION_FEES
                Set frmEdit = frmEditValuation
    
            Case BASE_RATES
                Set frmEdit = frmEditBaseRates
                
            '   AW  14/05/02    BM088
            Case INCOME_MULTIPLE
                Set frmEdit = frmEditIncMultiples
                
            '   AW  14/05/02    BM087
            Case MP_MIG_RATE_SETS
                Set frmEdit = frmEditMPMigRates
                
            '   AW  14/05/02    BM017
            Case REDEM_FEE_SETS
                Set frmEdit = frmEditRedemFeeSets
            
            Case GLOBAL_PARAM_FIXED
                 Set frmEdit = frmEditGlobalFixed
    
            Case GLOBAL_PARAM_BANDED
                 Set frmEdit = frmEditGlobalBanded
    
            Case COUNTRIES
                 Set frmEdit = frmEditCountry
            
            Case COMBOBOX_ENTRIES
                 Set frmEdit = frmEditCombo
    
            Case ERROR_MESSAGES
                 Set frmEdit = frmErrors
     
            Case DEPARTMENTS
                 Set frmEdit = frmEditDepartment
            
            Case REGIONS
                 Set frmEdit = frmEditRegion
            
            Case UNITS
                 Set frmEdit = frmEditUnit
    
            Case COMPETENCIES
                 Set frmEdit = frmEditCompetency
    
            Case WORKING_HOURS
                 Set frmEdit = frmEditWorkingHoursType
            
            Case DISTRIBUTION_CHANNELS
                 Set frmEdit = frmEditDistChannel
    
            Case LIFE_COVER_RATES
                 Set frmEdit = frmEditLifeCover
            
            Case BUILDINGS_AND_CONTENTS_PRODUCTS
                 Set frmEdit = frmEditBuildingsAndContents
    
            Case PAYMENT_PROTECTION_RATES
                 Set frmEdit = frmEditPayProtRates
    
            Case PAYMENT_PROTECTION_PRODUCTS
                 Set frmEdit = frmEditPayProtProducts
            
            Case MORTGAGE_PRODUCTS
                HandleProducts LV_PRODUCTS_EDIT, frmMain.lvListView
            
            ' --- DJP Phase 2 Task Management Start
            Case TASK_MANAGEMENT_TASKS
                Set frmEdit = frmEditTasks
            
            Case TASK_MANAGEMENT_STAGES
                Set frmEdit = frmEditStages
            
            Case TASK_MANAGEMENT_ACTIVITIES
            ' --- DJP Phase 2 Task Management End
                Set frmEdit = frmEditActivities
            
            Case LENDERS
                HandleLenders LV_LENDERS_EDIT
            
' TW 30/01/2007 EP2_1122
'            Case LOCAL_ADDRESS, PANEL_ADDRESS, LEGAL_REP_ADDRESS, VALUER_ADDRESS
            Case PANEL_ADDRESS, LEGAL_REP_ADDRESS, VALUER_ADDRESS
                'TK 29/11/2005 MAR81
                If frmMain.lvListView.ListItems.Item(frmMain.lvListView.SelectedItem.Index).SubItems(5) = "Inactive" Then
                    HandleActivateThirtParty
                Else
                    HandleEditThirtParty
                End If
            Case LOCAL_ADDRESS
                HandleEditThirtParty
' TW 30/01/2007 EP2_1122 End
            
            Case USERS
                HandleUsers LV_USER_EDIT, UserDetails
            
            Case APPLICATION_PROCESSING
                HandleAppProcessing LV_CANCEL_APPLICATION
                bCloseFind = False
            
            Case INCOME_FACTORS
                Set frmEdit = frmEditIncomeFactors
            
            Case ADDITIONAL_QUESTIONS
                Set frmEdit = frmEditQuestions
            Case BUSINESS_GROUPS
                Set frmEdit = frmEditBusinessGroup
            Case CONDITIONS
                Set frmEdit = frmEditConditions
            
            Case BATCH_SCHEDULER
                Set frmEdit = frmEditBatch
            
' TW 27/11/2007 DBM594
            Case PAYMENTS_FOR_COMPLETION
                Set frmEdit = frmEditPaymentForCompletions
' TW 27/11/2007 DBM594 End
            
            Case PRINTING_TEMPLATE
                Set frmEdit = frmEditPrintingTemplate
            'MAR45 GHun
            Case PRINTING_DOCUMENT
                Set frmEdit = frmEditPrintingDocument
            'MAR45 End
            
            'MAR202 Start
            Case PRINTING_PACK
                Set frmEdit = frmEditPrintingPack
            'MAR202 End
            
            Case BASE_RATE
                Set frmEdit = frmEditRate
            Case CURRENCIES
                Set frmEdit = frmEditCurrencies
            Case LEADAGENT, INTERMEDIARY_COMPANY, INDIVIDUAL, ADMIN_CENTRE, PACKAGER, BROKER ' EP15 pct ' EP522 PB
                Set frmEdit = frmEditIntermediaries
' TW 17/10/2006 EP2_15
            Case ASSOCIATIONS, CLUBS
                Set frmEdit = frmEditAssociation
                
            Case PACKAGERS, PRINCIPALS
                Set frmEdit = frmEditPackager

            Case ARFIRMS
                Set frmEdit = frmEditBroker
                
            Case ARBROKERS
                Set frmEdit = frmEditIndividualARBroker
                
            Case DABROKERS
                Set frmEdit = frmEditIndividualDABroker
                
            Case INDIVIDUALS & "_" & PACKAGERS
                Set frmEdit = frmEditIndividualPackager
' TW 17/10/2006 EP2_15 End
            Case Else
                ' Client specific
                Dim clsListView As ListViewCS
                Set clsListView = New ListViewCS
                
                clsListView.ListEdit vKey, frmEdit
            End Select
        End If
        
        If bCloseFind Then
            CloseFindDialogs
        End If
        
        If Not frmEdit Is Nothing Then
            frmEdit.SetKeys colMatchValues
            frmEdit.SetIsEdit True
            
            g_clsDataAccess.BeginTrans
            EndWaitCursor
            frmEdit.Show vbModal, frmMain
                
            enumReturn = frmEdit.GetReturnCode()
            
            Unload frmEdit
            
            HandleReturn enumReturn
        End If
    
        SetSelectedItem frmMain.lvListView, m_nCurrentIndex, SET_SELECTION
    Else
        g_clsErrorHandling.RaiseError errGeneralError, "Unable to locate key for item from TreeView"
    End If
    
    EndWaitCursor
    Exit Sub
Failed:
    If Not frmEdit Is Nothing Then
        g_clsDataAccess.RollbackTrans
        Set frmEdit = Nothing
    End If
    
    g_clsDataAccess.CloseConnection
    g_clsErrorHandling.DisplayError
End Sub
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' Function      : GetListViewKeys
' Description   : Returns the keys for the row in the listview that is currently
'                 selected.
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Public Sub GetListViewKeys(lv As MSGListView, colKeyValues As Collection)
    On Error GoTo Failed
    Dim nListIndex As Integer
    Dim clsPopulateDetails As PopulateDetails
    
    If Not lv.SelectedItem Is Nothing Then
        nListIndex = lv.SelectedItem.Index
        Set clsPopulateDetails = lv.GetExtra(nListIndex)
        
        If Not clsPopulateDetails Is Nothing Then
        'Set colKeyValues = frmMain.lvListView.GetExtra(nListIndex)
            Set colKeyValues = clsPopulateDetails.GetKeyMatchValues()
        
            If colKeyValues.Count = 0 Then
                g_clsErrorHandling.RaiseError errGeneralError, "Unable to obtain keys"
            End If
        Else
            g_clsErrorHandling.RaiseError errGeneralError, " Populate Details object not set"
        End If
    Else
        g_clsErrorHandling.RaiseError errGeneralError, "No listview selection"
    End If
    Exit Sub
Failed:
    g_clsErrorHandling.RaiseError Err.Number, Err.DESCRIPTION
End Sub
Private Sub GetLockApplicationNumbers(colAppNos As Collection)
    On Error GoTo Failed
    Dim lstItem As ListItem
    
    For Each lstItem In frmMain.lvListView.ListItems
        If lstItem.Selected = True Then
            MsgBox lstItem.Text
        End If
    Next
    Exit Sub
Failed:
    g_clsErrorHandling.RaiseError Err.Number, Err.DESCRIPTION
End Sub

Friend Sub GetLTVAmountAdjustments(lvListView As MSGListView)
    On Error GoTo Failed
    Dim clsProcFeeAdjByLTV As ProcFeeAdjByLTVTable
    Dim clsTableAccess As TableAccess
    
    Set clsProcFeeAdjByLTV = New ProcFeeAdjByLTVTable
    Set clsTableAccess = clsProcFeeAdjByLTV
    
    clsProcFeeAdjByLTV.GetLTVAmountAdjustments
    
    ' Populate the listview
    g_clsFormProcessing.PopulateFromRecordset lvListView, clsTableAccess
    
    Exit Sub
Failed:
    g_clsErrorHandling.RaiseError Err.Number, Err.DESCRIPTION
End Sub
Friend Sub GetLoanAmountAdjustments(lvListView As MSGListView)
    On Error GoTo Failed
    Dim clsProcFeeAdjByLoan As ProcFeeAdjByLoanTable
    Dim clsTableAccess As TableAccess
    
    Set clsProcFeeAdjByLoan = New ProcFeeAdjByLoanTable
    Set clsTableAccess = clsProcFeeAdjByLoan
    
    clsProcFeeAdjByLoan.GetLoanAmountAdjustments
    
    ' Populate the listview
    g_clsFormProcessing.PopulateFromRecordset lvListView, clsTableAccess
    
    Exit Sub
Failed:
    g_clsErrorHandling.RaiseError Err.Number, Err.DESCRIPTION
End Sub

Friend Sub GetDefaultProcurationFees(lvListView As MSGListView)
    On Error GoTo Failed
    Dim clsProcFeeDefault As ProcFeeDefaultTable
    Dim clsTableAccess As TableAccess
    
    Set clsProcFeeDefault = New ProcFeeDefaultTable
    Set clsTableAccess = clsProcFeeDefault
    
    clsProcFeeDefault.GetDefaultProcurationFees
    
    ' Populate the listview
    g_clsFormProcessing.PopulateFromRecordset lvListView, clsTableAccess
    
    Exit Sub
Failed:
    g_clsErrorHandling.RaiseError Err.Number, Err.DESCRIPTION
End Sub

Friend Sub GetFirmPackagers(lvListView As MSGListView)
' TW 16/01/2007 EP2_859
Dim strSQL As String
Dim clsPrincipalFirm As PrincipalFirmTable
Dim clsTableAccess As TableAccess

    On Error GoTo Failed
    
' TW 01/02/2007 EP2_1036
'    strSQL = "Select PRINCIPALFIRMID, FSARef, UNITID, PrincipalFirmName, " & _
'                "REPLACE(REPLACE(LTRIM(" & _
'                        "ISNULL(ADDRESSLINE1, '') + ' ' + ISNULL(ADDRESSLINE2, '') + ' ' + ISNULL(ADDRESSLINE3, '') + ' ' + ISNULL(ADDRESSLINE4, '') + ' ' + ISNULL(ADDRESSLINE5, '') + ' ' + ISNULL(ADDRESSLINE6, '') + ' ' + " & _
'                        "ISNULL(POSTCODE, '')) , '  ', ' '), '  ', ' ') AS ADDRESS, " & _
'                        "POSTCODE " & _
'                        "From PRINCIPALFIRM " & _
'                        "WHERE PACKAGERINDICATOR = 1"

    strSQL = "Select PRINCIPALFIRMID, UNITID, PrincipalFirmName, " & _
                "REPLACE(REPLACE(LTRIM(" & _
                        "ISNULL(ADDRESSLINE1, '') + ' ' + ISNULL(ADDRESSLINE2, '') + ' ' + ISNULL(ADDRESSLINE3, '') + ' ' + ISNULL(ADDRESSLINE4, '') + ' ' + ISNULL(ADDRESSLINE5, '') + ' ' + ISNULL(ADDRESSLINE6, '') + ' ' + " & _
                        "ISNULL(POSTCODE, '')) , '  ', ' '), '  ', ' ') AS ADDRESS, " & _
                        "POSTCODE " & _
                        "From PRINCIPALFIRM " & _
                        "WHERE PACKAGERINDICATOR = 1"
' TW 01/02/2007 EP2_1036 End

    
    Set clsPrincipalFirm = New PrincipalFirmTable
    Set clsTableAccess = clsPrincipalFirm
    
    clsTableAccess.GetTableDataFromSearch strSQL
    
    ' Populate the listview
    g_clsFormProcessing.PopulateFromRecordset lvListView, clsTableAccess
    
    Exit Sub
Failed:
    g_clsErrorHandling.RaiseError Err.Number, Err.DESCRIPTION
End Sub
Friend Sub GetARFirms(lvListView As MSGListView)
' TW 16/01/2007 EP2_859
Dim strSQL As String
Dim clsARFirm As ARFirmTable
Dim clsTableAccess As TableAccess

    On Error GoTo Failed
    
    strSQL = "Select ARFIRMID, FSAARFirmRef, ARFirmName, UnitId, " & _
                "REPLACE(REPLACE(LTRIM(" & _
                        "ISNULL(ADDRESSLINE1, '') + ' ' + ISNULL(ADDRESSLINE2, '') + ' ' + ISNULL(ADDRESSLINE3, '') + ' ' + ISNULL(ADDRESSLINE4, '') + ' ' + ISNULL(ADDRESSLINE5, '') + ' ' + ISNULL(ADDRESSLINE6, '') + ' ' + " & _
                        "ISNULL(POSTCODE, '')) , '  ', ' '), '  ', ' ') AS ADDRESS, " & _
                        "POSTCODE " & _
                        "From ARFirm"

    
    Set clsARFirm = New ARFirmTable
    Set clsTableAccess = clsARFirm
    
    clsTableAccess.GetTableDataFromSearch strSQL
    
    ' Populate the listview
    g_clsFormProcessing.PopulateFromRecordset lvListView, clsTableAccess
    
    Exit Sub
Failed:
    g_clsErrorHandling.RaiseError Err.Number, Err.DESCRIPTION
End Sub
Friend Sub GetAssociations(lvListView As MSGListView)
' TW 16/01/2007 EP2_859
Dim strSQL As String
Dim clsMortgageClubNetAssocTable As MortgageClubNetAssocTable
Dim clsTableAccess As TableAccess

    On Error GoTo Failed
    
    strSQL = "SELECT CLUBNETWORKASSOCIATIONID, MORTGAGECLUBNETWORKASSOCNAME, " & _
              "REPLACE(REPLACE(LTRIM(" & _
              "ISNULL(BUILDINGORHOUSENAME, '') + ' ' + ISNULL(BUILDINGORHOUSENUMBER, '') + ' ' + ISNULL(FLATNUMBER, '') + ' ' + ISNULL(STREET, '') + ' ' + ISNULL(DISTRICT, '') + ' ' + ISNULL(TOWN, '') + ' ' + ISNULL(COUNTY, '') + ' ' + " & _
              "ISNULL(POSTCODE, '')) , '  ', ' '), '  ', ' ') AS ADDRESS " & _
              "FROM " & _
              "MORTGAGECLUBNETWORKASSOCIATION " & _
              "WHERE PACKAGERINDICATOR = 1"
    
    Set clsMortgageClubNetAssocTable = New MortgageClubNetAssocTable
    Set clsTableAccess = clsMortgageClubNetAssocTable
    
    clsTableAccess.GetTableDataFromSearch strSQL
    
    ' Populate the listview
    g_clsFormProcessing.PopulateFromRecordset lvListView, clsTableAccess
    
    Exit Sub
Failed:
    g_clsErrorHandling.RaiseError Err.Number, Err.DESCRIPTION
End Sub
Friend Sub GetClubs(lvListView As MSGListView)
' TW 16/01/2007 EP2_859
Dim strSQL As String
Dim clsMortgageClubNetAssocTable As MortgageClubNetAssocTable
Dim clsTableAccess As TableAccess

    On Error GoTo Failed
    
    strSQL = "SELECT CLUBNETWORKASSOCIATIONID, MORTGAGECLUBNETWORKASSOCNAME, " & _
              "REPLACE(REPLACE(LTRIM(" & _
              "ISNULL(BUILDINGORHOUSENAME, '') + ' ' + ISNULL(BUILDINGORHOUSENUMBER, '') + ' ' + ISNULL(FLATNUMBER, '') + ' ' + ISNULL(STREET, '') + ' ' + ISNULL(DISTRICT, '') + ' ' + ISNULL(TOWN, '') + ' ' + ISNULL(COUNTY, '') + ' ' + " & _
              "ISNULL(POSTCODE, '')) , '  ', ' '), '  ', ' ') AS ADDRESS " & _
              "FROM " & _
              "MORTGAGECLUBNETWORKASSOCIATION " & _
              "WHERE ISNULL(PACKAGERINDICATOR, 0) = 0"
    
    Set clsMortgageClubNetAssocTable = New MortgageClubNetAssocTable
    Set clsTableAccess = clsMortgageClubNetAssocTable
    
    clsTableAccess.GetTableDataFromSearch strSQL
    
    ' Populate the listview
    g_clsFormProcessing.PopulateFromRecordset lvListView, clsTableAccess
    
    Exit Sub
Failed:
    g_clsErrorHandling.RaiseError Err.Number, Err.DESCRIPTION
End Sub
Friend Sub GetPrincipals(lvListView As MSGListView)
' TW 16/01/2007 EP2_859
Dim strSQL As String
Dim clsPrincipalFirm As PrincipalFirmTable
Dim clsTableAccess As TableAccess

    On Error GoTo Failed
    
    strSQL = "Select PRINCIPALFIRMID, FSARef, UNITID, PrincipalFirmName, " & _
                "REPLACE(REPLACE(LTRIM(" & _
                        "ISNULL(ADDRESSLINE1, '') + ' ' + ISNULL(ADDRESSLINE2, '') + ' ' + ISNULL(ADDRESSLINE3, '') + ' ' + ISNULL(ADDRESSLINE4, '') + ' ' + ISNULL(ADDRESSLINE5, '') + ' ' + ISNULL(ADDRESSLINE6, '') + ' ' + " & _
                        "ISNULL(POSTCODE, '')) , '  ', ' '), '  ', ' ') AS ADDRESS, " & _
                        "POSTCODE " & _
                        "From PRINCIPALFIRM " & _
                        "WHERE ISNULL(PACKAGERINDICATOR, 0) = 0"

    
    Set clsPrincipalFirm = New PrincipalFirmTable
    Set clsTableAccess = clsPrincipalFirm
    
    clsTableAccess.GetTableDataFromSearch strSQL
    
    ' Populate the listview
    g_clsFormProcessing.PopulateFromRecordset lvListView, clsTableAccess
    
    Exit Sub
Failed:
    g_clsErrorHandling.RaiseError Err.Number, Err.DESCRIPTION
End Sub

Friend Sub GetIndividualPackagers(lvListView As MSGListView)
    On Error GoTo Failed
    Dim clsIntroducerPackager As IntroducerPackagerTable
    Dim clsTableAccess As TableAccess
    
    Set clsIntroducerPackager = New IntroducerPackagerTable
    Set clsTableAccess = clsIntroducerPackager
    
    clsIntroducerPackager.GetIndividualPackagers
    
    ' Populate the listview
    g_clsFormProcessing.PopulateFromRecordset lvListView, clsTableAccess
    
    Exit Sub
Failed:
    g_clsErrorHandling.RaiseError Err.Number, Err.DESCRIPTION
End Sub

Friend Sub GetIndividualARBrokers(lvListView As MSGListView)
    On Error GoTo Failed
    Dim clsIntroducerARIndividual As IntroducerARFirmTable
    Dim clsTableAccess As TableAccess
    
    Set clsIntroducerARIndividual = New IntroducerARFirmTable
    Set clsTableAccess = clsIntroducerARIndividual
    
    clsIntroducerARIndividual.GetIndividualARBrokers
    
    ' Populate the listview
    g_clsFormProcessing.PopulateFromRecordset lvListView, clsTableAccess
    
    Exit Sub
Failed:
    g_clsErrorHandling.RaiseError Err.Number, Err.DESCRIPTION
End Sub

Friend Sub GetIndividualDABrokers(lvListView As MSGListView)
    On Error GoTo Failed
    Dim clsIntroducerDAFirm As IntroducerDAFirmTable
    Dim clsTableAccess As TableAccess
    
    Set clsIntroducerDAFirm = New IntroducerDAFirmTable
    Set clsTableAccess = clsIntroducerDAFirm
    
    clsIntroducerDAFirm.GetIndividualDABrokers
    
    ' Populate the listview
    g_clsFormProcessing.PopulateFromRecordset lvListView, clsTableAccess
    
    Exit Sub
Failed:
    g_clsErrorHandling.RaiseError Err.Number, Err.DESCRIPTION
End Sub

''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' Function      : GetOmigaUsers
' Description   : Populates the listview passed in with the Users from the
'                 Omiga User table
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Friend Sub GetOmigaUsers(lvListView As MSGListView, intDetail As LV_DETAIL)
    On Error GoTo Failed
    Dim clsOmigaUser As OmigaUserTable
    Dim clsTableAccess As TableAccess
    
    Set clsOmigaUser = New OmigaUserTable
    Set clsTableAccess = clsOmigaUser
    
    clsOmigaUser.GetUsers "", "", "", intDetail
    
    ' Populate the listview
    g_clsFormProcessing.PopulateFromRecordset lvListView, clsTableAccess
    
    Exit Sub
Failed:
    g_clsErrorHandling.RaiseError Err.Number, Err.DESCRIPTION
End Sub

''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' Function      : GetPaymentProtectionRates
' Description   : Populates the listview passed in with the payment protection
'                 rates.
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Friend Sub GetPaymentProtectionRates(lvListView As MSGListView)
    Dim clsPayProt As New PayProtRatesTable
    Dim clsTableAccess As TableAccess
    
    clsPayProt.GetPaymentProtectionRates

    Set clsTableAccess = clsPayProt
    
    ' Populate the listview
    g_clsFormProcessing.PopulateFromRecordset lvListView, clsTableAccess
    
End Sub
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' Function      : GetLifeCoverRates
' Description   : Populates the listview passed in with the Life Cover
'                 rates.
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Friend Sub GetLifeCoverRates(lvListView As MSGListView)
    On Error GoTo Failed
    Dim clsLifeCover As New LifeCoverRatesTable
        
    clsLifeCover.GetLifeCoverRates
    g_clsFormProcessing.PopulateFromRecordset lvListView, clsLifeCover
    
    Exit Sub
Failed:
    g_clsErrorHandling.RaiseError Err.Number, Err.DESCRIPTION
End Sub
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' Function      : GetCompetencies
' Description   : Populates the listview passed in with the Competencies
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Friend Sub GetCompetencies(lvListView As MSGListView)
    On Error GoTo Failed
    Dim clsCompetency As New CompetencyTable
        
    clsCompetency.GetCompetencies
    
    g_clsFormProcessing.PopulateFromRecordset lvListView, clsCompetency
    Exit Sub
Failed:
    g_clsErrorHandling.RaiseError Err.Number, Err.DESCRIPTION
End Sub
Friend Sub GetRegions(lvListView As MSGListView)
    On Error GoTo Failed
    Dim clsRegion As New RegionTable

    clsRegion.GetRegions

    g_clsFormProcessing.PopulateFromRecordset lvListView, clsRegion

    Exit Sub
Failed:
    g_clsErrorHandling.RaiseError Err.Number, Err.DESCRIPTION
End Sub
Friend Sub GetPanelAddress(lvListView As MSGListView, enumPanelType As ThirdPartyCombo)
    On Error GoTo Failed
    Dim clsPanel As New PanelTable

    clsPanel.GetPanelValues enumPanelType

    g_clsFormProcessing.PopulateFromRecordset lvListView, clsPanel

    Exit Sub
Failed:
    g_clsErrorHandling.RaiseError Err.Number, Err.DESCRIPTION
End Sub
Friend Sub GetLocalAddress(lvListView As MSGListView)
    On Error GoTo Failed
    Dim clsLocal As New NameAndAddressDirTable

    clsLocal.GetNamesAndAddresses
    g_clsFormProcessing.PopulateFromRecordset lvListView, clsLocal

    Exit Sub
Failed:
    g_clsErrorHandling.RaiseError Err.Number, Err.DESCRIPTION
End Sub
Friend Sub GetLenders(lvListView As MSGListView)
    On Error GoTo Failed
    Dim clsLender As New MortgageLendersTable

    clsLender.GetLenders
    g_clsFormProcessing.PopulateFromRecordset lvListView, clsLender

    Exit Sub
Failed:
    g_clsErrorHandling.RaiseError Err.Number, Err.DESCRIPTION
End Sub
Friend Sub GetActiveApplications(lvListView As MSGListView, Optional bUseExistingSearch As Boolean = False)
    On Error GoTo Failed
    Dim clsApp As ApplicationTable
    Set clsApp = New ApplicationTable
    
    clsApp.GetActiveApplications , , , bUseExistingSearch
    g_clsFormProcessing.PopulateFromRecordset lvListView, clsApp
    
    Exit Sub
Failed:
    g_clsErrorHandling.RaiseError Err.Number, Err.DESCRIPTION
End Sub
Friend Sub GetProducts(lvListView As MSGListView)
    On Error GoTo Failed
    Dim clsMortProd As New MortgageProductTable

    clsMortProd.GetProducts
    g_clsFormProcessing.PopulateFromRecordset lvListView, clsMortProd

    Exit Sub
Failed:
    g_clsErrorHandling.RaiseError Err.Number, Err.DESCRIPTION
End Sub
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' Function      : GetStages
' Description   : Populates the listview passed in with the Task Management Stages
' DJP Phase 2 Task Management
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Friend Sub GetStages(lvListView As MSGListView)
    On Error GoTo Failed
    
    Dim clsStages As StageTable
    
    Set clsStages = New StageTable
    clsStages.GetStagesWithCombos

    ' Populate the listview
    g_clsFormProcessing.PopulateFromRecordset lvListView, TableAccess(clsStages)
    
    Exit Sub
Failed:
    g_clsErrorHandling.RaiseError Err.Number, Err.DESCRIPTION
End Sub
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' Function      : GetTasks
' Description   : Populates the listview passed in with the Task Management Tasks
' DJP Phase 2 Task Management
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Friend Sub GetTasks(lvListView As MSGListView)
    On Error GoTo Failed
    
    Dim clsTasks As TaskTable
    
    Set clsTasks = New TaskTable
    clsTasks.GetTasksWithCombos

    ' Populate the listview
    g_clsFormProcessing.PopulateFromRecordset lvListView, TableAccess(clsTasks)
    
    Exit Sub
Failed:
    g_clsErrorHandling.RaiseError Err.Number, Err.DESCRIPTION
End Sub
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' Function      : GetActivities
' Description   : Populates the listview passed in with the Task Management Activities
' DJP Phase 2 Task Management
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Friend Sub GetActivities(lvListView As MSGListView)
    On Error GoTo Failed
    Dim clsActivities As ActivityTable
    
    Set clsActivities = New ActivityTable
    clsActivities.GetActivities
    
    ' Populate the listview
    g_clsFormProcessing.PopulateFromRecordset lvListView, TableAccess(clsActivities)
    
    Exit Sub
Failed:
    g_clsErrorHandling.RaiseError Err.Number, Err.DESCRIPTION
End Sub
Friend Sub GetApplicationLocks(lvListView As MSGListView, nAppType As ApplicationType)
    On Error GoTo Failed
    Dim clsAppLock As New ApplicationLockTable
    
    clsAppLock.GetApplicationLocks nAppType
    g_clsFormProcessing.PopulateFromRecordset lvListView, clsAppLock
    
    Exit Sub
Failed:
    g_clsErrorHandling.RaiseError Err.Number, Err.DESCRIPTION
End Sub
Friend Sub GetCustomerLocks(lvListView As MSGListView, nAppType As ApplicationType)
    On Error GoTo Failed
    Dim clsCustomerLock As New CustomerLockTable
    
    clsCustomerLock.GetCustomerLocks nAppType
    
    g_clsFormProcessing.PopulateFromRecordset lvListView, clsCustomerLock
    frmMain.lvListView.MultiSelect = True
    Exit Sub
Failed:
    g_clsErrorHandling.RaiseError Err.Number, Err.DESCRIPTION
End Sub

' ik_bm0314
Friend Sub GetDocumentLocks(lvListView As MSGListView)
    On Error GoTo Failed
    Dim clsDocumentLocks As New DocumentLock
    
    clsDocumentLocks.GetDocumentLocks
    
    g_clsFormProcessing.PopulateFromRecordset lvListView, clsDocumentLocks
    frmMain.lvListView.MultiSelect = True
    Exit Sub
Failed:
    g_clsErrorHandling.RaiseError Err.Number, Err.DESCRIPTION
End Sub

''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' Function      : GetIncomeFactors
' Description   : Populate the main listview with all defined income factors.
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Friend Sub GetIncomeFactors(ByRef lvListView As MSGListView)
        
    Dim clsIncomeFactorTable As AllowableIncomeFactorTable
    
    On Error GoTo Failed
    
    Set clsIncomeFactorTable = New AllowableIncomeFactorTable
    clsIncomeFactorTable.GetIncomeFactors
    
    ' Populate the listview
    g_clsFormProcessing.PopulateFromRecordset lvListView, TableAccess(clsIncomeFactorTable)
    
    Exit Sub
    
Failed:
    g_clsErrorHandling.RaiseError Err.Number, Err.DESCRIPTION
End Sub
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' Function      : GetData
' Description   : Populates the listview passed in with the data obtained from the
'                 table class passed in.
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Friend Sub GetData(clsTableAccess As TableAccess, lvListView As MSGListView, Optional enumPopulateType As PopulateType = POPULATE_KEYS, Optional enumClassOption As ClassOption)
    On Error GoTo Failed
    Dim rs As ADODB.Recordset
    Set rs = clsTableAccess.GetTableData(enumPopulateType, enumClassOption)
    'PB 19/10/2006 EP2_13 Begin
    If Not g_blnDoNotValidateRS Then
        ValidateRecordset rs, clsTableAccess.GetTable()
    Else
        g_blnDoNotValidateRS = False
    End If
    'PB 19/10/2006 EP2-13 End
    g_clsFormProcessing.PopulateFromRecordset lvListView, clsTableAccess
    
    clsTableAccess.CloseRecordSet
    Exit Sub
Failed:
    g_clsErrorHandling.RaiseError Err.Number, Err.DESCRIPTION
End Sub
Friend Sub SetUpdated(Optional bUpdated As Boolean = True)
    m_bUpdated = bUpdated
End Sub
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' Function      : HandleUpdates
' Description   : When an edit has taken place, this method decides what to do
'                 regarding the listview, either to repopulate it or not.
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Friend Sub HandleUpdates()
    On Error GoTo Failed
    Dim vKey As Variant
    
    ' After an update, we don't always want to repopulate the listview, so check first the item we have updated
    ' warrants a refresh
    
    vKey = frmMain.GetSelectedTreeKey()
    
    If (Not IsNull(vKey)) Then
        Select Case vKey
            ' These items do not want a repopulate of the list view
            ' DJP BM0282
' TW 03/02/2007 EP2_1101
'            Case LOCAL_ADDRESS, VALUER_ADDRESS, LEGAL_REP_ADDRESS, MORTGAGE_PRODUCTS, USERS, CONDITIONS
            Case LOCAL_ADDRESS, VALUER_ADDRESS, LEGAL_REP_ADDRESS, MORTGAGE_PRODUCTS, USERS
' TW 03/02/2007 EP2_1101 End
            Dim sCaption As String
            'BS BM0282 25/03/03
            'sCaption = frmMain.lblTitle(constListViewLabel).Caption
            sCaption = frmMain.tvwDB.SelectedItem()
            frmMain.lblTitle(constListViewLabel).Caption = sCaption & " (List needs refreshing)"
        Case Else
            ' But all others do...
            frmMain.PopulateListView
        End Select
    End If
    
    Exit Sub
Failed:
    g_clsErrorHandling.RaiseError Err.Number, Err.DESCRIPTION
End Sub
Private Sub Class_Initialize()
    m_bLogging = False
    Set m_clsIntermediary = New Intermediary

    m_bUpdated = False
    m_nCurrentIndex = -1
    
    supervisorObjects.bTaskManagement = False
End Sub
Friend Sub PromoteRecord()
    On Error GoTo Failed
    Dim colMatchValues As Collection
    Dim nResponse As Integer
    Dim vKey As Variant
    Dim lstItem As ListItem
    Dim clsPopulateDetails As PopulateDetails
    Dim sDescription As String
' TW 23/03/2007 EP2_1942
    Dim intPromotedCount As Integer
' TW 23/03/2007 EP2_1942 End
    
    If Not frmMain.lvListView.SelectedItem Is Nothing Then
        nResponse = MsgBox("Promote the selected record(s)?", vbQuestion + vbYesNo)

        If nResponse = vbYes Then
            GetListViewKeys frmMain.lvListView, colMatchValues

                ' Now decide what it is we have to delete
            vKey = frmMain.GetSelectedTreeKey()

            BeginWaitCursor

            If (Not IsNull(vKey)) Then
                For Each lstItem In frmMain.lvListView.ListItems
                    If lstItem.Selected = True Then
                        
                        Set clsPopulateDetails = frmMain.lvListView.GetExtra(lstItem.Index)
                        Set colMatchValues = clsPopulateDetails.GetKeyMatchValues()
                        
                        sDescription = clsPopulateDetails.GetObjectDescription()
                        
' TW 23/03/2007 EP2_1942 - This code is designed to force any blocks of
' records marked for promotion at the same time to have a different time
' of change allocated to each record. This prevents the situation where
' records 'disappear' from the list of records available for promotion.
'                        g_clsHandleUpdates.SaveChangeRequestFromName CStr(vKey), colMatchValues, sDescription
                        g_clsHandleUpdates.SaveChangeRequestFromName CStr(vKey), colMatchValues, sDescription, intPromotedCount
                        intPromotedCount = intPromotedCount + 1
' TW 23/03/2007 EP2_1942 End
                    End If
                Next
            End If
            EndWaitCursor
        End If
    End If
    
    Exit Sub
Failed:
    g_clsErrorHandling.RaiseError Err.Number, Err.DESCRIPTION
End Sub
Private Function CanObjectBeDeleted(sObjectName As String) As Boolean
    On Error GoTo Failed
    Dim bCanBeDeleted     As Boolean
    Dim clsDeleteObjectCS As DeleteObjectCS
    bCanBeDeleted = True

    Select Case sObjectName
        Case COMBOBOX_ENTRIES
' TW 09/10/2006 EP2_7
        Case ADDITIONAL_BORROWING_FEES
        Case CREDIT_LIMIT_INCREASE_FEES
' TW 09/10/2006 EP2_7 End
' TW 11/12/2006 EP2_20
        Case TRANSFER_OF_EQUITY_FEES
' TW 11/12/2006 EP2_20 End
' TW 14/12/2006 EP2_518
        Case DEFAULT_PROCURATION_FEES
        Case LOAN_AMOUNT_ADJUSTMENTS
        Case LTV_AMOUNT_ADJUSTMENTS
' TW 14/12/2006 EP2_518 End
        Case ADMIN_FEES
        '*=[MC]BMIDS763 - FEE SETS
        Case PRODUCT_SWITCH_FEESETS
        Case INSURANCE_ADMIN_FEESETS
        Case TT_FEESETS
        '*=[MC]SECTION END - BMIDS763
        Case VALUATION_FEES
        Case BASE_RATES
        Case RENTAL_INCOME_RATES  ' JD BMIDS765
        Case INCOME_MULTIPLE
        Case MP_MIG_RATE_SETS
        Case REDEM_FEE_SETS
        Case GLOBAL_PARAM_FIXED
        Case GLOBAL_PARAM_BANDED
        Case MORTGAGE_PRODUCTS
        Case COUNTRIES
        Case LIFE_COVER_RATES
        Case BUILDINGS_AND_CONTENTS_PRODUCTS
        Case PAYMENT_PROTECTION_RATES
        Case PAYMENT_PROTECTION_PRODUCTS
        Case LOCAL_ADDRESS, PANEL_ADDRESS, LEGAL_REP_ADDRESS, VALUER_ADDRESS
        Case USERS
        Case ERROR_MESSAGES
        Case LOCKS_ONLINE_APPLICATION, LOCKS_OFFLINE_APPLICATION
        Case LOCKS_ONLINE_CUSTOMER, LOCKS_OFFLINE_CUSTOMER
        Case TASK_MANAGEMENT_STAGES, TASK_MANAGEMENT_TASKS, TASK_MANAGEMENT_ACTIVITIES
        Case ADDITIONAL_QUESTIONS
        Case BUSINESS_GROUPS
        Case CONDITIONS
        Case BATCH_SCHEDULER
        Case BASE_RATE
        Case CURRENCIES
        Case INCOME_FACTORS
        Case Else
            Set clsDeleteObjectCS = New DeleteObjectCS
            bCanBeDeleted = clsDeleteObjectCS.CanObjectBeDeleted(sObjectName)

    End Select

    CanObjectBeDeleted = bCanBeDeleted
    
    Exit Function
Failed:
    g_clsErrorHandling.RaiseError Err.Number, Err.DESCRIPTION
End Function
Friend Function DeleteObject(sObjectName As String, _
                             colMatchValues As Collection, _
                             Optional sTargetDatabase As String, _
                             Optional bRecordChange As Boolean = True, _
                             Optional sObjectDescription As String) As Boolean
    On Error GoTo Failed
    Dim bObjectDeleted     As Boolean
    Dim clsTableAccess As TableAccess
    bObjectDeleted = True

    Set clsTableAccess = g_clsHandleUpdates.GetObjectTableClass(sObjectName, False)
    clsTableAccess.SetDatabase sTargetDatabase

    Select Case sObjectName
    Case LOCKS_ONLINE_APPLICATION, LOCKS_OFFLINE_APPLICATION
        Dim clsAppLock As ApplicationLockTable
        Set clsAppLock = clsTableAccess
        clsAppLock.DeleteApplicationAndCustomers colMatchValues(1)
        bRecordChange = False
    
    Case LOCKS_ONLINE_CUSTOMER, LOCKS_OFFLINE_CUSTOMER
        Dim clsCustLock As CustomerLockTable
        Set clsCustLock = clsTableAccess
        clsCustLock.DeleteCustomer colMatchValues(1)
        bRecordChange = False
    
    ' ik_bm0314
    Case LOCKS_DOCS
        Dim clsDocLock As DocumentLock
        Set clsDocLock = clsTableAccess
        clsTableAccess.SetKeyMatchValues colMatchValues
        clsDocLock.DeleteDocumentLock colMatchValues
        bRecordChange = False

    Case COMBOBOX_ENTRIES
        Dim clsComboGroup As ComboValueGroupTable
        Set clsComboGroup = clsTableAccess
        
        clsTableAccess.SetKeyMatchValues colMatchValues
        clsComboGroup.DeleteComboGroup sTargetDatabase
' TW 09/10/2006 EP2_7
    Case CREDIT_LIMIT_INCREASE_FEES
        clsTableAccess.SetKeyMatchValues colMatchValues
        clsTableAccess.DeleteRecords

        'BM00023 - Delete the parent set record if all bands have been deleted.
        Dim clsCreditLimitIncreaseFeeTable As CreditLimitIncFeeSetTable
        Set clsCreditLimitIncreaseFeeTable = New CreditLimitIncFeeSetTable
        clsCreditLimitIncreaseFeeTable.DeleteIfNoBands colMatchValues(1), sTargetDatabase

    Case ADDITIONAL_BORROWING_FEES
        clsTableAccess.SetKeyMatchValues colMatchValues
        clsTableAccess.DeleteRecords

        'BM00023 - Delete the parent set record if all bands have been deleted.
        Dim clsAdditionalBorrowingFeeTable As AddnlBorrFeeSetTable
        Set clsAdditionalBorrowingFeeTable = New AddnlBorrFeeSetTable
        clsAdditionalBorrowingFeeTable.DeleteIfNoBands colMatchValues(1), sTargetDatabase
' TW 09/10/2006 EP2_7 End
' TW 11/12/2006 EP2_20
    Case TRANSFER_OF_EQUITY_FEES
        clsTableAccess.SetKeyMatchValues colMatchValues
        clsTableAccess.DeleteRecords

        'BM00023 - Delete the parent set record if all bands have been deleted.
        Dim clsTransferOfEquityFeeTable As TransferOfEquityFeeSetTable
        Set clsTransferOfEquityFeeTable = New TransferOfEquityFeeSetTable
        clsTransferOfEquityFeeTable.DeleteIfNoBands colMatchValues(1), sTargetDatabase
' TW 11/12/2006 EP2_20 End
' TW 14/12/2006 EP2_518
    Case DEFAULT_PROCURATION_FEES
        clsTableAccess.SetKeyMatchValues colMatchValues
        clsTableAccess.DeleteRecords
    Case LOAN_AMOUNT_ADJUSTMENTS
        clsTableAccess.SetKeyMatchValues colMatchValues
        clsTableAccess.DeleteRecords
    Case LTV_AMOUNT_ADJUSTMENTS
        clsTableAccess.SetKeyMatchValues colMatchValues
        clsTableAccess.DeleteRecords
' TW 14/12/2006 EP2_518 End

    Case ADMIN_FEES
        clsTableAccess.SetKeyMatchValues colMatchValues
        clsTableAccess.DeleteRecords

        'BM00023 - Delete the parent set record if all bands have been deleted.
        Dim clsAdminFeeTable As AdminFeeSetTable
        Set clsAdminFeeTable = New AdminFeeSetTable
        clsAdminFeeTable.DeleteIfNoBands colMatchValues(1), sTargetDatabase
    '*=[MC]BMIDS763 - CC075 - FEESETS
    Case PRODUCT_SWITCH_FEESETS
        clsTableAccess.SetKeyMatchValues colMatchValues
        clsTableAccess.DeleteRecords
        
        Dim objProductSwictchFeeTable As ProductSwitchFeeSetTable
        Set objProductSwictchFeeTable = New ProductSwitchFeeSetTable
        
        objProductSwictchFeeTable.DeleteIfNoBands colMatchValues(1), sTargetDatabase
    Case INSURANCE_ADMIN_FEESETS
        clsTableAccess.SetKeyMatchValues colMatchValues
        clsTableAccess.DeleteRecords
        
        Dim objInsuFeeSetTable As InsuranceAdminFeeSetTable
        Set objInsuFeeSetTable = New InsuranceAdminFeeSetTable
        
        objInsuFeeSetTable.DeleteIfNoBands colMatchValues(1), sTargetDatabase
        
    Case TT_FEESETS
        clsTableAccess.SetKeyMatchValues colMatchValues
        clsTableAccess.DeleteRecords
        
        Dim objTTFeeSetTable As TTFeeSetTable
        Set objTTFeeSetTable = New TTFeeSetTable
        
        objTTFeeSetTable.DeleteIfNoBands colMatchValues(1), sTargetDatabase
        
    '*=[MC]BMIDS763 - CC075 - FEESETS SECTION END
    Case VALUATION_FEES
        clsTableAccess.SetKeyMatchValues colMatchValues
        clsTableAccess.DeleteRecords
        
        'BM00023 - Delete the parent set record if all bands have been deleted.
        Dim clsValFeeTable As ValuationFeeSetTable
        Set clsValFeeTable = New ValuationFeeSetTable
        clsValFeeTable.DeleteIfNoBands colMatchValues(1), sTargetDatabase

    Case BASE_RATES
        clsTableAccess.SetKeyMatchValues colMatchValues
        clsTableAccess.DeleteRecords
        
        'SYS4247 - Delete the parent set record if all bands have been deleted.
        Dim clsRateSetTable As BaseRateSetTable
        Set clsRateSetTable = New BaseRateSetTable
        clsRateSetTable.DeleteIfNoBands colMatchValues(1)
        'SYS4247 - End.

    Case RENTAL_INCOME_RATES  'JD BMIDS765
        clsTableAccess.SetKeyMatchValues colMatchValues
        clsTableAccess.DeleteRecords
        
        'Delete the parent set record if all bands have been deleted.
        Dim clsRentalIncomeRateSetTable As RentalIncomeRateSetTable
        Set clsRentalIncomeRateSetTable = New RentalIncomeRateSetTable
        clsRentalIncomeRateSetTable.DeleteIfNoBands colMatchValues(1)
        
    Case INCOME_MULTIPLE
        clsTableAccess.SetKeyMatchValues colMatchValues
        clsTableAccess.DeleteRecords
       
    Case MP_MIG_RATE_SETS
        clsTableAccess.SetKeyMatchValues colMatchValues
        clsTableAccess.DeleteRecords
        
        Dim clsMigSetTable As MPMigRateSetTable
        Set clsMigSetTable = New MPMigRateSetTable
        clsMigSetTable.DeleteIfNoBands colMatchValues(1), sTargetDatabase
        
     Case REDEM_FEE_SETS
        clsTableAccess.SetKeyMatchValues colMatchValues
        clsTableAccess.DeleteRecords

        Dim clsRedemFeeTable As RedemptionFeeSetTable
        Set clsRedemFeeTable = New RedemptionFeeSetTable
        clsRedemFeeTable.DeleteIfNoBands colMatchValues(1), sTargetDatabase
     
    Case GLOBAL_PARAM_FIXED
        clsTableAccess.SetKeyMatchValues colMatchValues
        clsTableAccess.DeleteRecords

    Case GLOBAL_PARAM_BANDED
        clsTableAccess.SetKeyMatchValues colMatchValues
        clsTableAccess.DeleteRecords

    Case MORTGAGE_PRODUCTS
        Dim clsMortProd As MortgageProductTable
        Set clsMortProd = clsTableAccess
        
        clsTableAccess.SetKeyMatchValues colMatchValues
        clsMortProd.DeleteProduct

    Case COUNTRIES
        Dim clsCountry As CountryTable
        Set clsCountry = clsTableAccess
        
        clsTableAccess.SetKeyMatchValues colMatchValues
        clsCountry.DeleteCountry

    Case LIFE_COVER_RATES
        clsTableAccess.SetKeyMatchValues colMatchValues
        clsTableAccess.DeleteRecords

    Case BUILDINGS_AND_CONTENTS_PRODUCTS
        clsTableAccess.SetKeyMatchValues colMatchValues
        clsTableAccess.DeleteRecords

    Case PAYMENT_PROTECTION_RATES
        clsTableAccess.SetKeyMatchValues colMatchValues
        clsTableAccess.DeleteRecords

    Case PAYMENT_PROTECTION_PRODUCTS
        clsTableAccess.SetKeyMatchValues colMatchValues
        clsTableAccess.DeleteRecords

    Case LOCAL_ADDRESS, PANEL_ADDRESS, LEGAL_REP_ADDRESS, VALUER_ADDRESS, NAMES_AND_ADDRESSES
        Dim clsPanel As PanelTable
        Set clsPanel = clsTableAccess
        
        clsTableAccess.SetKeyMatchValues colMatchValues
        clsPanel.DeletePanel

    Case USERS
        Dim clsUser As OmigaUserTable
        Set clsUser = clsTableAccess
        
        clsTableAccess.SetKeyMatchValues colMatchValues
        clsUser.DeleteUser

    Case ERROR_MESSAGES
        clsTableAccess.SetKeyMatchValues colMatchValues
        clsTableAccess.DeleteRecords
    
    Case TASK_MANAGEMENT_TASKS
        Dim clsTask As TaskTable
        Dim colMatchFields As Collection
        Dim clsTaskPriority As TaskPriority
        Dim clsStageTaskTable As StageTaskTable
        Dim clsLinkedTaskTable As LinkedTaskTable   'MAR1300 GHun
        
        'Delete Task Priority
        Set clsTaskPriority = New TaskPriority
        TableAccess(clsTaskPriority).SetKeyMatchValues colMatchValues
        TableAccess(clsTaskPriority).SetDatabase sTargetDatabase
        TableAccess(clsTaskPriority).GetTableData POPULATE_KEYS
        TableAccess(clsTaskPriority).DeleteRecords
        
        'Delete Linked tasks
        Set clsLinkedTaskTable = New LinkedTaskTable
        TableAccess(clsLinkedTaskTable).SetDatabase sTargetDatabase
        clsLinkedTaskTable.DeleteLinkedTasksForTaskId (colMatchValues.Item(1))
        
        'Set the delete flag on the StageTask table.
        Set clsStageTaskTable = New StageTaskTable
        Set colMatchFields = New Collection
        colMatchFields.Add "TASKID"
        TableAccess(clsStageTaskTable).SetKeyMatchFields colMatchFields
        TableAccess(clsStageTaskTable).SetKeyMatchValues colMatchValues
        TableAccess(clsStageTaskTable).GetTableData POPULATE_KEYS
        'CL SYS4707
        If TableAccess(clsStageTaskTable).RecordCount <> 0 Then
            clsStageTaskTable.SetDeleteFlag True
            TableAccess(clsStageTaskTable).Update
        End If
        'END SYS4707
                
        
        'Now Delete the task
        clsTableAccess.SetDatabase sTargetDatabase
        clsTableAccess.SetKeyMatchValues colMatchValues
        Set clsTask = clsTableAccess

        clsTask.DeleteTask
    
    Case TASK_MANAGEMENT_STAGES
        Dim clsStage As StageTable
        Dim clsStageTask As StageTaskTable
        
        Set clsStage = clsTableAccess
        Set clsStageTask = New StageTaskTable
        'First StageTask
        
        TableAccess(clsStageTask).SetDatabase sTargetDatabase
        clsStageTask.DeleteStageTaskForStage CStr(colMatchValues(1))
        
        clsTableAccess.SetKeyMatchValues colMatchValues
        clsTableAccess.SetDatabase sTargetDatabase
        clsStage.DeleteStage
    
    Case TASK_MANAGEMENT_ACTIVITIES
        Dim clsActivity As ActivityTable
        Set clsActivity = clsTableAccess
        
        clsTableAccess.SetKeyMatchValues colMatchValues
        clsActivity.DeleteActivity sTargetDatabase
        
    Case BUSINESS_GROUPS
        Dim clsTaskLink As New TaskLinkTable
        Dim clsGroup As New BusinessGroupTable
        
        Set clsTableAccess = clsTaskLink
        
        'Delete TaskLink Records
        clsTableAccess.SetKeyMatchValues colMatchValues
        clsTableAccess.DeleteRecords sTargetDatabase
        
        'Now the Group itself
        Set clsTableAccess = clsGroup
        clsTableAccess.SetKeyMatchValues colMatchValues
        clsTableAccess.DeleteRecords sTargetDatabase
        
    Case ADDITIONAL_QUESTIONS
        Dim clsQuestions As AdditionalQuestionsTable
        Set clsQuestions = clsTableAccess
        
        clsTableAccess.SetKeyMatchValues colMatchValues
        clsQuestions.DeleteQuestion
    
    Case PRINTING_TEMPLATE
        Dim clsPrint As New PrintingTemplateTable
        Dim clsAvailableTemplates As New AvailableTemplatesTable
        
        'Delete Stages
        TableAccess(clsAvailableTemplates).SetKeyMatchValues colMatchValues
        TableAccess(clsAvailableTemplates).DeleteRecords
        
        'Delete The Printing Template
        TableAccess(clsPrint).SetKeyMatchValues colMatchValues
        TableAccess(clsPrint).DeleteRecords
        bRecordChange = False
    Case CONDITIONS
        Dim clsConditions As ConditionsTable
        Set clsConditions = clsTableAccess
        
        clsTableAccess.SetKeyMatchValues colMatchValues
        clsConditions.DeleteCondition
        
    Case BATCH_SCHEDULER
        Dim clsBatch As BatchTable
        Dim clsPopulateDetails As PopulateDetails
        Dim sStatusType As String
        Set clsBatch = clsTableAccess
        
        Set clsPopulateDetails = frmMain.lvListView.GetExtra(frmMain.lvListView.SelectedItem.Index)
        sStatusType = clsPopulateDetails.GetExtra()
        
        clsTableAccess.SetKeyMatchValues colMatchValues
        clsBatch.DeleteBatch CLng(sStatusType)
        bRecordChange = False
    
    Case CURRENCIES
        clsTableAccess.SetKeyMatchValues colMatchValues
        clsTableAccess.DeleteRecords
        
    Case INCOME_FACTORS
        clsTableAccess.SetKeyMatchValues colMatchValues
        clsTableAccess.DeleteRecords

' TW 18/11/2006 EP2_132
    Case INTRODUCERFIRM
        clsTableAccess.SetKeyMatchValues colMatchValues
        clsTableAccess.DeleteRecords
' TW 18/11/2006 EP2_132 End
' TW 23/11/2006 EP2_172
    Case EXCLUSIVEPRODUCTS
        clsTableAccess.SetKeyMatchValues colMatchValues
        clsTableAccess.DeleteRecords
' TW 23/11/2006 EP2_172 End
' TW 30/11/2006 EP2_253
    Case MORTGAGEPRODUCTPRODUCTCLASS
        clsTableAccess.SetKeyMatchValues colMatchValues
        clsTableAccess.DeleteRecords
    Case MORTGAGEPRODUCTNATUREOFLOAN
        clsTableAccess.SetKeyMatchValues colMatchValues
        clsTableAccess.DeleteRecords
    Case MORTGAGEPRODUCTINCOMESTATUS
        clsTableAccess.SetKeyMatchValues colMatchValues
        clsTableAccess.DeleteRecords
' TW 30/11/2006 EP2_253 End
' TW 06/12/2006 EP2_330
    Case FIRMPERMISSIONS
        clsTableAccess.SetKeyMatchValues colMatchValues
        clsTableAccess.DeleteRecords
    Case NETWORKASSOCIATIONS
        clsTableAccess.SetKeyMatchValues colMatchValues
        clsTableAccess.DeleteRecords
    Case FIRMTRADINGNAME
        clsTableAccess.SetKeyMatchValues colMatchValues
        clsTableAccess.DeleteRecords
' TW 06/12/2006 EP2_330 End
' TW 05/02/2007 EP2_706
    Case APPOINTMENTS
        clsTableAccess.SetKeyMatchValues colMatchValues
        clsTableAccess.DeleteRecords
' TW 05/02/2007 EP2_706 End
    Case Else
        Dim clsDeleteObjectCS As DeleteObjectCS
        Set clsDeleteObjectCS = New DeleteObjectCS
        
        bObjectDeleted = clsDeleteObjectCS.DeleteObject(sObjectName, clsTableAccess, colMatchValues, sTargetDatabase, bRecordChange, sObjectDescription)
        
        If Not bObjectDeleted Then
            g_clsErrorHandling.RaiseError errGeneralError, "Supervisor does not support the deletion of this item"
        End If
        
    End Select

    If bObjectDeleted And bRecordChange Then
        g_clsHandleUpdates.SaveChangeRequest clsTableAccess, sObjectDescription, PromoteDelete
    End If
    
    DeleteObject = bObjectDeleted
    Exit Function
Failed:
    g_clsErrorHandling.RaiseError Err.Number, Err.DESCRIPTION
End Function
Friend Sub DeleteRecord()
    Dim nResponse As Integer
    Dim vKey As Variant
    Dim sObjectDescription As String
    Dim colMatchValues As New Collection
    Dim bObjectDeleted As Boolean
    Dim lstItem As ListItem
    Dim clsPopulateDetails As PopulateDetails
    
    On Error GoTo Failed
    bObjectDeleted = True
    
    ' Now decide what it is we have to delete
    vKey = frmMain.GetSelectedTreeKey()
    
    If CanObjectBeDeleted(CStr(vKey)) Then
        If Not frmMain.lvListView.SelectedItem Is Nothing Then
            bObjectDeleted = True
            nResponse = MsgBox("Delete the selected item(s)?", vbQuestion + vbYesNo)
        
            If nResponse = vbYes Then
                GetListViewKeys frmMain.lvListView, colMatchValues
                    
                
                BeginWaitCursor
                
                If (Not IsNull(vKey)) Then
                    g_clsDataAccess.BeginTrans
                    
                    For Each lstItem In frmMain.lvListView.ListItems
                        If lstItem.Selected = True Then
                            Set clsPopulateDetails = frmMain.lvListView.GetExtra(lstItem.Index)
 
                            Set colMatchValues = clsPopulateDetails.GetKeyMatchValues()
                            sObjectDescription = clsPopulateDetails.GetObjectDescription()
                            bObjectDeleted = DeleteObject(CStr(vKey), colMatchValues, , , sObjectDescription)
                        
                            If Not bObjectDeleted Then
                                Exit For
                            End If
                        End If
                    Next
                    
                    If bObjectDeleted Then
                        g_clsDataAccess.CommitTrans
                        HandleUpdates
                        g_clsMainSupport.SetSelectedItem frmMain.lvListView, m_nCurrentIndex, CLEAR_SELECTION
                    
    
                        ' Record the fact that we've deleted this item
                    Else
                        g_clsDataAccess.RollbackTrans
                    End If
                End If
            End If
        End If
    Else
        MsgBox "Supervisor does not support the deletion of '" + CStr(vKey) + "'", vbCritical
    End If
    
    EndWaitCursor
    Exit Sub
Failed:
    g_clsErrorHandling.DisplayError "Unable to delete item - check it isn't in use by Omiga 4."
    g_clsDataAccess.RollbackTrans
    
    Exit Sub
    Resume
End Sub
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' Function      : CreateProduct
' Description   : Called to create a Mortgage Product.
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' Function      : CreateProduct
' Description   : Called to create a Mortgage Product.
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Function CreateProduct() As MSGReturnCode
    On Error GoTo Failed
    Dim enumReturn As MSGReturnCode
    Dim clsMortgageProduct As MortgageProduct
    
    ' Create the Product class
    Set clsMortgageProduct = New MortgageProductCS

    clsMortgageProduct.SetIsEdit False
    
    ' Start a transaction - this will be committed or rolled back when the form exits.
    g_clsDataAccess.BeginTrans
    
    ' Show the form
    clsMortgageProduct.Show
    
    ' Get the return code and pass back, but first unload the form
    enumReturn = clsMortgageProduct.GetReturnCode
    clsMortgageProduct.UnloadForm

    CreateProduct = enumReturn
    
    Exit Function
Failed:
    g_clsErrorHandling.RaiseError Err.Number, Err.DESCRIPTION
End Function
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' Function      : CreateLender
' Description   : Called to create a Lender
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Function CreateLender() As MSGReturnCode
    On Error GoTo Failed
    Dim clsLenderCS As Lender
    Dim enumReturn As MSGReturnCode
    
    ' Create the Lender class
    Set clsLenderCS = New LenderCS
    clsLenderCS.SetIsEdit False
    
    ' Start a transaction - this will be committed or rolled back when the form exits.
    g_clsDataAccess.BeginTrans
    
    ' Show the form
    clsLenderCS.Show
    
    ' Get the return code and pass back, but first unload the form
    enumReturn = clsLenderCS.GetReturnCode
    clsLenderCS.UnloadForm
    
    CreateLender = enumReturn
    
    Exit Function
Failed:
    g_clsErrorHandling.RaiseError Err.Number, Err.DESCRIPTION
End Function
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' Function      : HandleLenders
' Description   : This method is called when the user right clicks on the listview
'                 when Lenders are listed. The right click menu can show any one of
'                 of a number of tabs, so this method allows any of the tabs to be
'                 displayed.
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Friend Function HandleLenders(nIndex As Integer) As MSGReturnCode
    Dim activeTab As LenderDetailsTabs
    Dim nListIndex As Integer
    Dim colMatchValues As Collection
    Dim clsLenderCS As Lender
    Dim enumReturn As MSGReturnCode
    
    On Error GoTo Failed
    Set clsLenderCS = New LenderCS
    
    enumReturn = MSGFailure
    BeginWaitCursor
        
    activeTab = NoTab
    SetSelectedItem frmMain.lvListView, m_nCurrentIndex
    
    ' Find the tab we want to show
    Select Case nIndex
        Case LV_LENDERS_EDIT
            activeTab = LenderDetails
    
        Case LV_LENDERS_LEGAL_FEES
            activeTab = LegalFees
    
        Case LV_LENDERS_OTHER_FEES
            activeTab = OtherFees
        
        Case LV_LENDERS_MIG_RATES
            activeTab = MIGRateSets
        
        Case LV_LENDERS_CONTACT_DETAILS
            activeTab = ContactDetails
        
        Case LV_LENDERS_PARAMETERS
            activeTab = AdditionalParameters
        
        Case LV_LENDERS_PROMOTE
            PromoteRecord
        Case Else
            g_clsErrorHandling.RaiseError errGeneralError, "HandleLenders: Unknown index, " & nIndex
    End Select

    If activeTab <> NoTab Then
        Dim clsPopulateDetails As PopulateDetails
        nListIndex = frmMain.lvListView.SelectedItem.Index
                
        Set clsPopulateDetails = frmMain.lvListView.GetExtra(nListIndex)
        Set colMatchValues = clsPopulateDetails.GetKeyMatchValues
    
        If Not colMatchValues Is Nothing Then
            If colMatchValues.Count > 0 Then
                clsLenderCS.SetKeys colMatchValues
            Else
                g_clsErrorHandling.RaiseError errGeneralError, "Key values missing from listview for Lenders"
            End If
        Else
            g_clsErrorHandling.RaiseError errGeneralError, "Unable to obtain keys to match selection"
        End If
        
        ' Load the form
        g_clsDataAccess.BeginTrans
        clsLenderCS.SetIsEdit True
        clsLenderCS.SetActiveTab activeTab, False
        
        EndWaitCursor
        clsLenderCS.Show
    
        ' If it's done any updates, repopulate
        enumReturn = clsLenderCS.GetReturnCode()
        
        clsLenderCS.UnloadForm
        HandleReturn enumReturn
    End If
    
    SetSelectedItem frmMain.lvListView, m_nCurrentIndex, SET_SELECTION
    
    EndWaitCursor
    HandleLenders = enumReturn
    Exit Function
Failed:
    g_clsErrorHandling.SaveError
    g_clsDataAccess.RollbackTrans
    g_clsDataAccess.CloseConnection
    g_clsErrorHandling.RaiseError
End Function
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' Function      : HandleProducts
' Description   : This method is called when the user right clicks on the listview
'                 when products are listed. The right click menu can show any one of
'                 of a number of tabs, so this method allows any of the tabs to be
'                 displayed.
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Friend Function HandleProducts(nIndex As Integer, lvListView As MSGListView) As MSGReturnCode
    Dim nTab As ProductTabs
    Dim bLoadForm As Boolean
    Dim enumReturn As MSGReturnCode
    Dim colMatchValues As Collection
    Dim clsMortgageProduct As MortgageProduct
    On Error GoTo Failed
    
    enumReturn = MSGFailure
    
    BeginWaitCursor
    SetSelectedItem frmMain.lvListView, m_nCurrentIndex
    Set clsMortgageProduct = New MortgageProductCS
    
    bLoadForm = True
    ' Find the tab we need to display
    Select Case nIndex
    Case LV_PRODUCTS_COPY
        CopyRecord
        bLoadForm = False
    
    Case LV_PRODUCTS_DELETE
        DeleteRecord
        bLoadForm = False
    Case LV_PRODUCTS_EDIT
        nTab = mortProdDetails
    
' TW 22/02/2007 EP2_1577
'    Case LV_PRODUCTS_PURPOSE_OF_LOAN, LV_PRODUCTS_TYPE_OF_APP_ELIG, _
'         LV_PRODUCTS_PURPOSE_OF_LOAN, LV_PRODUCTS_TYPE_OF_BUYER_ELIG
'         nTab = mortProdAppEligibility
'
'    Case LV_PRODUCTS_INTEREST_RATES, LV_PRODUCTS_BASE_RATE_SETS
'        nTab = mortProdInteresRates
'
'    Case LV_PRODUCTS_SPECIAL_GROUPS, LV_PRODUCTS_TYPICAL_APR, LV_PRODUCTS_MIG_RATES, LV_PRODUCTS_INCOME_MULTIPLES
'        nTab = mortProdMisc
'
'    Case LV_PRODUCTS_EMPLOYMENT_ELIG, LV_PRODUCTS_CHANNEL_ELIG, _
'        LV_PRODUCTS_PROPERTY_LOCATION
'        nTab = mortProdOtherEligibility
'
'    Case LV_PRODUCTS_INCENTIVE_SETS, LV_PRODUCTS_OTHER_FEES, LV_PRODUCTS_REDEMPTION_FEES
'        nTab = mortProdOtherFeesAndIncentives
'
'     Case LV_PRODUCTS_SPECIAL_CONDITIONS
'        nTab = mortProdSpecialConditions
'
'    Case LV_PRODUCTS_ADMIN_FEE, LV_PRODUCTS_VALUATION_FEE
'        nTab = mortProdFees
'
'    Case LV_PRODUCTS_PARAMETERS_BANDS_SETS
'        nTab = mortProdParametersAndBands
' TW 22/02/2007 EP2_1577 End
    
    Case LV_PRODUCTS_PROMOTE
        bLoadForm = False
        PromoteRecord
    
    Case LV_PRODUCTS_VIEW_ERRORS
        bLoadForm = False
        ShowProductErrors
    
    Case Else
        g_clsErrorHandling.RaiseError errGeneralError, "Can't display tab (" & nIndex & ") doesn't exist"
    End Select

    ' Load the form
    If bLoadForm Then
        Set colMatchValues = New Collection
        
        g_clsFormProcessing.GetListViewKeys frmMain.lvListView, colMatchValues
        
        clsMortgageProduct.SetKeys colMatchValues
        
        clsMortgageProduct.SetIsEdit True
        clsMortgageProduct.SetActiveTab nTab
        g_clsDataAccess.BeginTrans
        
        clsMortgageProduct.Show
        
        enumReturn = clsMortgageProduct.GetReturnCode
        clsMortgageProduct.UnloadForm
        
        HandleReturn enumReturn

    End If
    
    SetSelectedItem frmMain.lvListView, m_nCurrentIndex, SET_SELECTION
    EndWaitCursor
    
    HandleProducts = enumReturn
    
    Exit Function
Failed:
    g_clsErrorHandling.SaveError
    g_clsDataAccess.RollbackTrans
    g_clsErrorHandling.RaiseError
End Function
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' Function      : IsProductValid
' Description   : Returns TRUE if the product is valid, false if not.
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Public Function IsProductValid() As Boolean
    On Error GoTo Failed
    Dim sValid As String
    Dim bValid As Boolean
    Dim colProduct As Collection
    Dim clsProductDetails As PopulateDetails
    
    bValid = True
    
    If g_clsVersion.DoesVersioningExist() = False Then
        Set colProduct = frmMain.lvListView.GetLine(, clsProductDetails)
        
        If Not clsProductDetails Is Nothing Then
            sValid = clsProductDetails.GetExtra()
        
            If Len(sValid) > 0 Then
                bValid = CBool(sValid)
            End If
        End If
    End If
    
    IsProductValid = bValid
    
    Exit Function
Failed:
    g_clsErrorHandling.RaiseError Err.Number, Err.DESCRIPTION
End Function
' Added by PJO 28/11/2005 MAR81
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' Function      : IsLegalRepInactive
' Description   : Returns TRUE if the product is valid, false if not.
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Public Function IsLegalRepInactive() As Boolean
    On Error GoTo Failed
    Dim sStatus As String
    Dim bInactive As Boolean
    
    bInactive = True
    Dim clsPanelLegalRep As New PanelLegalRepTable
'    sStatus = clsPanelLegalRep.GetLegalRepStatus()
    sStatus = "10"
    If sStatus = "10" Then
        bInactive = False
    End If
    IsLegalRepInactive = bInactive
    
    Exit Function
Failed:
    g_clsErrorHandling.RaiseError Err.Number, Err.DESCRIPTION
End Function
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' Function      : ShowProductErrors
' Description   : Dispalys a list of errors associated with the Mortgage Product currently selected
'                 on the main listview.
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Public Sub ShowProductErrors()
    On Error GoTo Failed
    Dim colProduct As Collection
    Dim clsProductDetails As PopulateDetails
    
    Set colProduct = frmMain.lvListView.GetLine(, clsProductDetails)
    
    If Not clsProductDetails Is Nothing Then
        Dim colMatchValues As Collection
        Set colMatchValues = clsProductDetails.GetKeyMatchValues
    
        frmProductErrors.SetKeyMatchValues colMatchValues
        frmProductErrors.Show vbModal, frmMain
        
        Unload frmProductErrors
    End If
    
    Exit Sub
Failed:
    g_clsErrorHandling.RaiseError Err.Number, Err.DESCRIPTION
End Sub
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' Function      : HandleReturn
' Description   : When any forms are displayed, a transaction is started. When the form returns, the
'                 transaction is commited or rolled back depending on the return code. This function
'                 will handle committing the transaction or rolling it back.
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Public Sub HandleReturn(enumReturn As MSGReturnCode, Optional bPopulateListView As Boolean = True)
    On Error GoTo Failed
    If enumReturn = MSGSuccess Then
        g_clsDataAccess.CommitTrans
     Else
        g_clsDataAccess.RollbackTrans
    End If
    
    ' DJP BM0318 Only repopulate if this method is asked to, and if the previous action was successful
    If bPopulateListView And enumReturn = MSGSuccess Then
        'Refresh the listview regardless of the last transaction.
        HandleUpdates
    End If
    
    ' Close the connection as we don't need it anymore.
    g_clsDataAccess.CloseConnection
    
    Exit Sub
Failed:
    g_clsErrorHandling.RaiseError Err.Number, Err.DESCRIPTION
End Sub
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' Function      : RemoveLock
' Description   : Removes a Customer or Application Lock. The lock to be removed will be the currently
'                 selected lock.
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Public Sub RemoveLock(enumLockType As LockType)
    On Error GoTo Failed
    'Dim vAppNo As Variant
    Dim clsLock As New LockProcessing
    Dim colKeys As New Collection
    Dim lstItem As ListItem
    Dim clsPopulateDetails As PopulateDetails
    
    BeginWaitCursor
    
    ' May be more than one lock selected so loop through looking for items that are selected.
    For Each lstItem In frmMain.lvListView.ListItems
        If lstItem.Selected = True Then
        Set clsPopulateDetails = frmMain.lvListView.GetExtra(lstItem.Index)
        Set colKeys = clsPopulateDetails.GetKeyMatchValues()
            
' ik_bm0314
'            Select Case enumLockType
'
'            Case Application
'                vAppNo = colKeys(1)
'
'                clsLock.SetLockType enumLockType
'                clsLock.RemoveLock vAppNo
'
'            Case Customer
'                vAppNo = colKeys(1)
'
'                clsLock.SetLockType enumLockType
'                clsLock.RemoveLock vAppNo
'
'            End Select
            
        clsLock.SetLockType enumLockType
        clsLock.RemoveLock colKeys
' ik_bm0314_ends
        
        End If
    Next
    
    EndWaitCursor
    Exit Sub
Failed:
    g_clsErrorHandling.RaiseError Err.Number, Err.DESCRIPTION
End Sub
Public Sub HandleBatchProcessing(nIndex As Integer)
    On Error GoTo Failed
    Dim colMatchValues As Collection
    Dim clsComboValidation As ComboValidationTable
    Dim frmShow As Form
    Dim bRet As Boolean
    Dim clsTableAccess As TableAccess
    Dim clsBatch As BatchTable
    Dim sComboGroup As String
    GetListViewKeys frmMain.lvListView, colMatchValues
    
    sComboGroup = "BatchStatus"
    Set clsComboValidation = New ComboValidationTable
    Set clsBatch = New BatchTable
    Set clsTableAccess = clsBatch

    Select Case nIndex
        Case LV_BATCH_PROCESS_VIEWJOBS
            Set frmShow = frmBatchSchedule
        Case LV_BATCH_PROCESS_VIEW
            Set frmShow = frmEditBatch
        Case LV_BATCH_PROCESS_DELETE
            bRet = DeleteObject(BATCH_SCHEDULER, colMatchValues)
        Case LV_BATCH_PROCESS_RESTART, LV_BATCH_PROCESS_LAUNCH, LV_BATCH_PROCESS_CANCEL
            clsTableAccess.SetKeyMatchValues colMatchValues
            bRet = clsBatch.UpdateBatchStatus(CLng(nIndex))
    End Select
    
    If Not frmShow Is Nothing Then
        frmShow.SetKeys colMatchValues
        frmShow.SetIsEdit True
        'SYS3327
        If nIndex = LV_BATCH_PROCESS_VIEW Then
            frmShow.SetIsView True
        End If
        frmShow.Show vbModal
        Unload frmShow
    End If
    
    If bRet Then
        HandleUpdates
    End If
    
    Set frmShow = Nothing
    Exit Sub
Failed:
    g_clsErrorHandling.RaiseError Err.Number, Err.DESCRIPTION
End Sub
Public Sub HandleAppProcessing(nIndex As Integer)
    On Error GoTo Failed
    Dim sAppNo As String
    Dim clsAppProcessing As ApplicationProcessing
    Dim colValues As Collection
    Dim nResponse As Integer
    ' Get the Application Number from the listview (i.e., the current selection)
    ' Only allowing one at a time for now.
    
    GetListViewKeys frmMain.lvListView, colValues
    
    ' Only one, and it's the app no, so get it
    sAppNo = CStr(colValues(1))
    
    Set clsAppProcessing = New ApplicationProcessing
    
    Select Case nIndex
        Case LV_CANCEL_APPLICATION
        nResponse = MsgBox("Cancel application ' " + sAppNo + "' ?", vbQuestion + vbYesNo)
                
        If nResponse = vbYes Then
            Dim bUseExistingSearch As Boolean
            clsAppProcessing.CancelApplication sAppNo
            bUseExistingSearch = True
            frmMain.PopulateListView bUseExistingSearch
            'HandleUpdates bUseExistingSearch
        End If
    End Select
    
    Exit Sub
Failed:
    g_clsErrorHandling.RaiseError Err.Number, Err.DESCRIPTION
End Sub
Public Sub HandleIntermediaries(nIndex As Integer)
    On Error GoTo Failed
    
    m_clsIntermediary.HandleIntermediaries nIndex

    
    Exit Sub
Failed:
    g_clsErrorHandling.RaiseError Err.Number, Err.DESCRIPTION
End Sub
Public Sub HandleLocks(nIndex As Integer, Optional bAskUser As Boolean = True)
    On Error GoTo Failed
    Dim vKey As Variant
    Dim enumLockType As LockType
    Dim clsTableAccess As TableAccess
    ' Which node was selected?
    vKey = frmMain.GetSelectedTreeKey()
    
    If (Not IsNull(vKey)) Then
        Select Case vKey
        
        Case LOCKS_ONLINE_APPLICATION, LOCKS_OFFLINE_APPLICATION
             enumLockType = Application
            
        Case LOCKS_ONLINE_CUSTOMER, LOCKS_OFFLINE_CUSTOMER
            Set clsTableAccess = New CustomerLockTable
            enumLockType = Customer

        ' ik_bm0314
        Case LOCKS_DOCS
            Set clsTableAccess = New DocumentLock
            enumLockType = Document

        Case Else
            g_clsErrorHandling.RaiseError errGeneralError, "Locks - Invalid tree selection " & vKey
        End Select
    
        Select Case nIndex
        
        Case LV_FIND_LOCK

' ik_bm0314
' is this necessary?
'            If enumLockType = Application Then
'                Set clsTableAccess = New ApplicationLockTable
'            Else
'                Set clsTableAccess = New CustomerLockTable
'            End If
            Select Case enumLockType
                Case Application
                    Set clsTableAccess = New ApplicationLockTable
                Case Customer
                    Set clsTableAccess = New CustomerLockTable
                Case Document
                    Set clsTableAccess = New DocumentLock
            End Select
' ik_bm0314_ends
            
            frmFindLocks.SetLockType enumLockType
            frmFindLocks.SetTableAccess clsTableAccess
            frmFindLocks.Show vbModal, frmMain

            'This is populated from the recordset, but it's not updated anymore.
            If frmFindLocks.GetReturnCode() = MSGSuccess Then
                If Not clsTableAccess Is Nothing Then
                    g_clsFormProcessing.PopulateFromRecordset frmMain.lvListView, clsTableAccess
                End If
            End If
        
            Unload frmFindLocks
        
        Case LV_CREATE_LOCK
            CreateLock enumLockType
        
        ' BM0282
        Case LV_ALL_LOCK
            frmMain.PopulateListView
        
        Case LV_REMOVE_LOCK
            Dim nResponse As Integer
            nResponse = vbYes
            
            If bAskUser Then
                nResponse = MsgBox("Delete the selected lock(s)?", vbQuestion + vbYesNo)
            End If
            
            If nResponse = vbYes Then
                g_clsMainSupport.RemoveLock enumLockType
                frmMain.PopulateListView
            End If
        
        End Select

    End If

    Exit Sub
Failed:
    g_clsErrorHandling.RaiseError Err.Number, Err.DESCRIPTION
End Sub
Public Sub CreateLock(enumLockType As LockType)
    On Error GoTo Failed
    
    frmCreateLock.SetLockType enumLockType
    frmCreateLock.Show vbModal, frmMain
    frmMain.PopulateListView
    Unload frmCreateLock
    
    Exit Sub
Failed:
    g_clsErrorHandling.RaiseError Err.Number, Err.DESCRIPTION
End Sub

Public Sub CopyRecord()
    On Error GoTo Failed
    Dim colMatchValues As Collection
    Dim vKey As Variant
    Dim enumReturn As MSGReturnCode
    Dim frmCopy As Form
    Set colMatchValues = New Collection
     
    vKey = frmMain.GetSelectedTreeKey()
    GetListViewKeys frmMain.lvListView, colMatchValues
    
    Select Case vKey
        Case MORTGAGE_PRODUCTS
        Set frmCopy = frmCopyProducts
    End Select
        
    If Not frmCopy Is Nothing Then
        frmCopy.SetKeys colMatchValues
        frmCopy.Show vbModal, frmMain
        enumReturn = frmCopy.GetReturnCode()
        Unload frmCopy
    End If
    
    If enumReturn = MSGSuccess Then
        HandleUpdates
    End If
    
    Exit Sub
Failed:
    g_clsErrorHandling.RaiseError Err.Number, Err.DESCRIPTION
End Sub
Private Sub HandleLogging()
    On Error GoTo Failed
    Dim sArg As String
    Dim nArgs As Long
    Dim nThisArg As Long
    Dim sCommandLine As Variant
    
    ' DJP - Added logging
    sCommandLine = GetCommandLine
    nArgs = UBound(sCommandLine)
    
    If nArgs > 0 Then
        For nThisArg = 1 To nArgs
            sArg = sCommandLine(nThisArg)
            
            If UCase(sArg) = ARG_LOGGING Then
                SetLogging
            End If
        Next
    End If
    
    
    Exit Sub
Failed:
    ' Not bothered what happens if this fails - the app should never stop because logging has failed.
End Sub
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' Function      : HandleStartup
' Description   : Performs all startup processing when Supervisor is started.
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Public Function HandleStartup() As Boolean
    On Error GoTo Failed
    Dim bRet As Boolean
    Dim clsConnections As DatabaseOptions

    
    HandleLogging
    
    ' DJP SQL Server port
    Set clsConnections = New DatabaseOptions
    
    ' Read in all database connections. Make sure:
    ' 1) There are some
    ' 2) There is an active one
    bRet = clsConnections.CreateConnections()
    
    HandleStartup = bRet
    
    Exit Function
Failed:
    g_clsErrorHandling.RaiseError Err.Number, Err.DESCRIPTION
End Function

Public Sub SetDatabaseCaption()
    Dim sDatabaseName As String
    
    sDatabaseName = g_clsDataAccess.GetActiveDatabaseName()
' TW 15/01/2007 EP2_826
'    frmMain.lblTitle(0).Caption = "Supervisor Functions (" & sDatabaseName & ")"
    frmMain.Caption = "Supervisor - Current Database - " & sDatabaseName
' TW 15/01/2007 EP2_826 End
End Sub
Public Sub EnableSupervisorObjects()
    On Error GoTo Failed
    Dim clsTaskManagement As TaskTable
    Dim bExists As Boolean
    
    Set clsTaskManagement = New TaskTable
    
    bExists = g_clsDataAccess.DoesTableExist(TableAccess(clsTaskManagement).GetTable())
    
    supervisorObjects.bTaskManagement = bExists
    
    ' Batch
    'BS BM0240 28/04/03
'#If ENABLE_BATCH = 1 Then
#If ENABLE_BATCH Then
    supervisorObjects.bBatch = True
#Else
    supervisorObjects.bBatch = False
#End If

    'Printing
    'BS BM0240 28/04/03
'#If ENABLE_PRINTING_TEMPLATE = 1 Then
#If ENABLE_PRINTING_TEMPLATE Then
    supervisorObjects.bPrintingTemplate = True
#Else
    supervisorObjects.bPrintingTemplate = False
#End If
    
    'Business Groups
    'BS BM0240 28/04/03
'#If ENABLE_BUSINESS_GROUPS = 1 Then
#If ENABLE_BUSINESS_GROUPS Then
    supervisorObjects.bBusinessGroups = True
#Else
    supervisorObjects.bBusinessGroups = False
#End If

    'Rate
    'BS BM0240 28/04/03
'#If ENABLE_RATE = 1 Then
#If ENABLE_RATE Then
    Dim clsIntermediary As IntermediaryTable
    Set clsIntermediary = New IntermediaryTable
    
    supervisorObjects.bBaseRate = True
#Else
    supervisorObjects.bBaseRate = False
#End If

    'Intermediary
    'BS BM0240 28/04/03
'#If ENABLE_INTERMEDIARY = 1 Then
#If ENABLE_INTERMEDIARY Then
    ' DJP SYS2564 if we say we want intermediaries, make sure the table is there first.
    supervisorObjects.bIntermediary = DoIntermediaryTablesExist
#Else
    supervisorObjects.bIntermediary = False
#End If

    ' DJP Record if Currencies exist or not.
    Dim clsCurrency As TableAccess
    Set clsCurrency = New CurrencyTable
    
    supervisorObjects.bCurrencies = g_clsDataAccess.DoesTableExist(clsCurrency.GetTable())
    
    Exit Sub
Failed:
    g_clsErrorHandling.RaiseError Err.Number, Err.DESCRIPTION
End Sub
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' Function      : DoIntermediaryTablesExist (SYS2564)
' Description   : Returns true if the Intermediary tables exist, false if not.
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Function DoIntermediaryTablesExist() As Boolean
    On Error GoTo Failed
    Dim sTable As String
    Dim bExists As Boolean
    Dim clsInt As IntermediaryTable
    Dim clsIntermediary As TableAccess

    Set clsIntermediary = New IntermediaryTable
    Set clsInt = clsIntermediary
    
    bExists = g_clsDataAccess.DoesTableExist(clsIntermediary.GetTable())

    sTable = clsInt.GetIndividualTable()
    bExists = g_clsDataAccess.DoesTableExist(sTable) And bExists
    
    sTable = clsInt.GetOrganisationTable()
    bExists = g_clsDataAccess.DoesTableExist(sTable) And bExists

    Set clsIntermediary = New IntBankAccountTable
    bExists = g_clsDataAccess.DoesTableExist(clsIntermediary.GetTable()) And bExists

    Set clsIntermediary = New IntCorrespondenceTable
    bExists = g_clsDataAccess.DoesTableExist(clsIntermediary.GetTable()) And bExists

    Set clsIntermediary = New IntCrossSellingTable
    bExists = g_clsDataAccess.DoesTableExist(clsIntermediary.GetTable()) And bExists

    Set clsIntermediary = New IntermediaryProcFeeTable
    bExists = g_clsDataAccess.DoesTableExist(clsIntermediary.GetTable()) And bExists

    Set clsIntermediary = New IntermediaryProdPeriodTable
    bExists = g_clsDataAccess.DoesTableExist(clsIntermediary.GetTable()) And bExists

    Set clsIntermediary = New IntermediaryReportDaysTable
    bExists = g_clsDataAccess.DoesTableExist(clsIntermediary.GetTable()) And bExists

    Set clsIntermediary = New IntProcFeeSplitForIntTable
    bExists = g_clsDataAccess.DoesTableExist(clsIntermediary.GetTable()) And bExists

    Set clsIntermediary = New IntProcFeeSplitTable
    bExists = g_clsDataAccess.DoesTableExist(clsIntermediary.GetTable()) And bExists

    DoIntermediaryTablesExist = bExists
    
    Exit Function
Failed:
    g_clsErrorHandling.RaiseError Err.Number, Err.DESCRIPTION
End Function

Public Function DoesRateExist() As Boolean
    DoesRateExist = supervisorObjects.bBaseRate
End Function
Public Function DoesBusinessGroupsExist() As Boolean
    DoesBusinessGroupsExist = supervisorObjects.bBusinessGroups
End Function
Public Function DoesIntermediaryExist() As Boolean
    DoesIntermediaryExist = supervisorObjects.bIntermediary
End Function
Public Function DoesPrintTemplateExist() As Boolean
    DoesPrintTemplateExist = supervisorObjects.bPrintingTemplate
End Function
Public Function DoesTaskManagementExist() As Boolean
    DoesTaskManagementExist = supervisorObjects.bTaskManagement
End Function
Public Function DoesBatchExist() As Boolean
    DoesBatchExist = supervisorObjects.bBatch
End Function
Public Function DoesCurrencyExist() As Boolean
    DoesCurrencyExist = supervisorObjects.bCurrencies
End Function
Public Function UserLogin(frmParent As Form, Optional bAllowDatabaseChange As Boolean = True) As Boolean
    On Error GoTo Failed
    Dim bSuccess As Boolean
    Dim clsUserTable As OmigaUserTable
            
    bSuccess = True
    'BS BM0282 28/04/03
    '#If login = 1 Then
    #If login Then

    If Not g_clsVersion.DoesVersioningExist() Then
        If Not g_bFirstTimeExecuted Then
            Dim nUserCount As Long
            Set clsUserTable = New OmigaUserTable
            
            ' DJP SYS4052 - only force the user to login if the target database actually has any
            ' users on it. If it doesn't, the database is blank so allow the user to create a base
            ' set or Organisation data.
            clsUserTable.GetUsers

            nUserCount = TableAccess(clsUserTable).RecordCount

            If nUserCount > 0 Then
                frmLogin.EnableChangeDatabase bAllowDatabaseChange
                frmLogin.Show vbModal, frmParent
                
                If Not frmLogin.GetReturnCode = MSGSuccess Then
                    bSuccess = False
                End If
                Unload frmLogin
            Else
                ' DJP SYS4093, if the user hasn't logged on, reset the UserID
                g_sSupervisorUser = ""
            End If
        End If
    End If
    #Else
        g_sSupervisorUser = ""
    #End If
    
    UserLogin = bSuccess
    Exit Function
Failed:
    g_clsErrorHandling.RaiseError Err.Number, Err.DESCRIPTION
End Function
Public Sub CloseFindDialogs()
    On Error GoTo Failed
    Unload frmFindApplication
    ' DJP BM0138 unload search screen if it is currently show (non modally)
    Unload frmFindPanel
    Exit Sub
Failed:
End Sub
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' Function      :   ShowSupervisorObjects
' Description   :   Decides which parts of Supervisor can and can't be used. At present, this is
'                   just enabling or disabling TaskManagement
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Public Sub ShowSupervisorObjects()
    On Error GoTo Failed
    
    ' Some of this may depend on whether or not certain tables exist, so clear the history first.
    g_clsDataAccess.ClearTableHistory
    
    ' Decide which can/cannot be shown
    EnableSupervisorObjects
    
    ' The following are methods that do things based on what happened above.
    ' DJP SQL Server port - clear the listview when creating the tree - we may have
    ' switched databases which will make the listview data possibly incorrect.
    frmMain.lvListView.ListItems.Clear
    frmMain.SetupTreeValues
    Exit Sub
Failed:
    g_clsErrorHandling.RaiseError Err.Number, Err.DESCRIPTION
End Sub

Friend Sub GetBusinessGroups(lvListView As MSGListView)
    On Error GoTo Failed
    
    Dim clsGroups As BusinessGroupTable
    
    Set clsGroups = New BusinessGroupTable
    
    clsGroups.GetGroups

    ' Populate the listview
    g_clsFormProcessing.PopulateFromRecordset lvListView, clsGroups
    
    Exit Sub
Failed:
    g_clsErrorHandling.RaiseError Err.Number, Err.DESCRIPTION
End Sub
Friend Sub GetQuestions(lvListView As MSGListView)
    On Error GoTo Failed
    
    Dim clsQuestion As AdditionalQuestionsTable
    
    Set clsQuestion = New AdditionalQuestionsTable
    
    clsQuestion.GetQuestions

    ' Populate the listview
    g_clsFormProcessing.PopulateFromRecordset lvListView, clsQuestion
    
    Exit Sub
Failed:
    g_clsErrorHandling.RaiseError Err.Number, Err.DESCRIPTION
End Sub

Friend Sub GetConditions(lvListView As MSGListView)
    On Error GoTo Failed
    
    Dim clsConditions As ConditionsTable
    
    Set clsConditions = New ConditionsTable
    
    clsConditions.GetConditions

    ' Populate the listview
    g_clsFormProcessing.PopulateFromRecordset lvListView, clsConditions
    
    Exit Sub
Failed:
    g_clsErrorHandling.RaiseError Err.Number, Err.DESCRIPTION
End Sub
'AA 05/03/01 - Printing Templates
Friend Sub GetPrintTemplates(lvListView As MSGListView)
    On Error GoTo Failed
    
    Dim clsPrinting As PrintingTemplateTable
    
    Set clsPrinting = New PrintingTemplateTable
    
    clsPrinting.GetPrintTemplates

    ' Populate the listview
    g_clsFormProcessing.PopulateFromRecordset lvListView, clsPrinting
    
    Exit Sub
Failed:
    g_clsErrorHandling.RaiseError Err.Number, Err.DESCRIPTION
End Sub

'MAR45 GHun
Friend Sub GetPrintDocuments(lvListView As MSGListView)
    On Error GoTo Failed
    
    Dim clsPrinting As TemplateTable
    
    Set clsPrinting = New TemplateTable
    
    TableAccess(clsPrinting).GetTableData POPULATE_ALL

    ' Populate the listview
    g_clsFormProcessing.PopulateFromRecordset lvListView, clsPrinting
    
    Exit Sub
Failed:
    g_clsErrorHandling.RaiseError Err.Number, Err.DESCRIPTION
End Sub
'MAR45 End

'MAR202 GHun
Friend Sub GetPrintPacks(lvListView As MSGListView)
    On Error GoTo Failed
    
    Dim clsPackControl As PackControlTable
    
    Set clsPackControl = New PackControlTable
    
    TableAccess(clsPackControl).GetTableData POPULATE_ALL

    ' Populate the listview
    g_clsFormProcessing.PopulateFromRecordset lvListView, clsPackControl
    
    Exit Sub
Failed:
    g_clsErrorHandling.RaiseError Err.Number, Err.DESCRIPTION
End Sub
'MAR202 End

Friend Sub GetIntermediaries(lvListView As MSGListView, clsTreeItem As TreeItem, tvNode As MSComctlLib.node)
    On Error GoTo Failed
    
    m_clsIntermediary.PopulateListView clsTreeItem, lvListView, tvNode
    
    Exit Sub
Failed:
    g_clsErrorHandling.RaiseError Err.Number, Err.DESCRIPTION
End Sub

'AA 13/02/01 - Batch Scheduler
Friend Sub GetBatch(lvListView As MSGListView)
    On Error GoTo Failed
    
    Dim clsBatch As BatchTable
    
    Set clsBatch = New BatchTable
    
    clsBatch.GetBatch

    ' Populate the listview
    g_clsFormProcessing.PopulateFromRecordset lvListView, clsBatch
    
    Exit Sub
Failed:
    g_clsErrorHandling.RaiseError Err.Number, Err.DESCRIPTION
End Sub
' TW 27/11/2007 DBM594
Friend Sub GetPaymentsForCompletion(lvListView As MSGListView)
    On Error GoTo Failed
    
    Dim clsDisbursementpaymentTable As DisbursementpaymentTable
    
    Set clsDisbursementpaymentTable = New DisbursementpaymentTable
    
    clsDisbursementpaymentTable.GetPaymentsForCompletion

    ' Populate the listview
    g_clsFormProcessing.PopulateFromRecordset lvListView, clsDisbursementpaymentTable
    
    Exit Sub
Failed:
    g_clsErrorHandling.RaiseError Err.Number, Err.DESCRIPTION
End Sub
' TW 27/11/2007 DBM594 End

'BM0425 No longer required
''AA 13/02/01 - Batch Schedule
'Friend Sub GetBatchSchedule(lvListView As MSGListView, sBatchNumber As String)
'    On Error GoTo Failed
'
'    Dim clsBatch As BatchScheduleTable
'
'    Set clsBatch = New BatchScheduleTable
'
'    clsBatch.GetBatchScheduleTableData sBatchNumber
'
'    ' Populate the listview
'    g_clsFormProcessing.PopulateFromRecordset lvListView, clsBatch
'
'    Exit Sub
'Failed:
'    g_clsErrorHandling.RaiseError Err.Number, Err.DESCRIPTION
'End Sub
'BM0425 End

'AA 26/02/01 - Base Rate
Friend Sub GetBaseRate(lvListView As MSGListView)
    On Error GoTo Failed
    
    Dim clsRate As RateTable
    
    Set clsRate = New RateTable
    
    clsRate.GetBaseRate

    ' Populate the listview
    g_clsFormProcessing.PopulateFromRecordset lvListView, clsRate
    
    Exit Sub
Failed:
    g_clsErrorHandling.RaiseError Err.Number, Err.DESCRIPTION
End Sub

Friend Sub GetRate(lvListView As MSGListView)
    On Error GoTo Failed
    
    Dim clsRate As RateTable
    
    Set clsRate = New RateTable
    
    clsRate.GetBaseRate

    ' Populate the listview
    g_clsFormProcessing.PopulateFromRecordset lvListView, clsRate
    
    Exit Sub
Failed:
    g_clsErrorHandling.RaiseError Err.Number, Err.DESCRIPTION
End Sub

Public Sub PopulateSingleIntermediary(tv As MSComctlLib.TreeView, node As MSComctlLib.node, clsTreeItem As TreeItem)
    On Error GoTo Failed
    
    m_clsIntermediary.PopulateSingleIntermediary tv, node, clsTreeItem
    
    Exit Sub
Failed:
    g_clsErrorHandling.RaiseError Err.Number, Err.DESCRIPTION

End Sub
Public Sub PopulateIntermediaries(tv As MSComctlLib.TreeView, node As MSComctlLib.node, sType As String)
    On Error GoTo Failed
        
    m_clsIntermediary.PopulateIntermediaries tv, node
        
    Exit Sub
Failed:
    g_clsErrorHandling.RaiseError Err.Number, Err.DESCRIPTION
End Sub
Friend Sub GetCurrencies(lvListView As MSGListView)
    On Error GoTo Failed
    
    Dim clsCurrency As CurrencyTable
    
    Set clsCurrency = New CurrencyTable
    
    clsCurrency.GetCurrencies

    ' Populate the listview
    g_clsFormProcessing.PopulateFromRecordset lvListView, clsCurrency
    
    Exit Sub
Failed:
    g_clsErrorHandling.RaiseError Err.Number, Err.DESCRIPTION
End Sub

Public Function GetIntermediary() As Intermediary
    Set GetIntermediary = m_clsIntermediary
End Function
Public Function GetIntermediaryTreeItem() As TreeItem
    Set GetIntermediaryTreeItem = m_clsIntermediary.GetTreeItem
End Function
' Pinched from VB Help
Function GetCommandLine(Optional ByVal MaxArgs As Integer)
   'Declare variables.
   Dim C As String, CmdLine As String, CmdLnLen As Long, InArg As Boolean, i As Long, NumArgs As Integer
   'See if MaxArgs was provided.
   If IsMissing(MaxArgs) Then MaxArgs = 10
   'Make array of the correct size.
   ReDim ArgArray(MaxArgs)
   NumArgs = 0: InArg = False
   'Get command line arguments.
   CmdLine = Command()
   CmdLnLen = Len(CmdLine)
   'Go thru command line one character
   'at a time.
   For i = 1 To CmdLnLen
      C = Mid(CmdLine, i, 1)
      'Test for space or tab.
      If (C <> " " And C <> vbTab) Then
         'Neither space nor tab.
         'Test if already in argument.
         If Not InArg Then
         'New argument begins.
         'Test for too many arguments.
            If NumArgs = MaxArgs Then Exit For
            NumArgs = NumArgs + 1
            InArg = True
         End If
         'Concatenate character to current argument.
         ArgArray(NumArgs) = ArgArray(NumArgs) & C
      Else
         'Found a space or tab.
         'Set InArg flag to False.
         InArg = False
      End If
   Next i
   'Resize array just enough to hold arguments.
   ReDim Preserve ArgArray(NumArgs)
   'Return Array in Function name.
   GetCommandLine = ArgArray()
End Function
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' Function      : SetLogging
' Description   : Starts logging for this application, if specified.
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Public Sub SetLogging(Optional bEnable As Boolean = True)
    On Error GoTo Failed
    m_bLogging = bEnable
    
    If m_bLogging Then
        App.StartLogging LOG_PATH, vbLogAuto
    End If
Failed:
End Sub
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' Function      : GetLogging
' Description   : Returns whether or not logging has been enabled for this application
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Public Function GetLogging() As Boolean
    On Error GoTo Failed
    
    GetLogging = m_bLogging
    Exit Function
Failed:
    GetLogging = False
End Function

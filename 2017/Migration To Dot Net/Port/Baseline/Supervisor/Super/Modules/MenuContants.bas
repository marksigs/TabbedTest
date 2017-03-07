Attribute VB_Name = "MenuContants"
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' Module        : MenuConstants.bas
' Description   : Constants for menu items.
' Change history
' Prog      Date        Description
' DJP       21/11/01    ' DJP       22/11/01    SYS2912 SQL Server locking problem.
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' Birmingham Midshires Change history
' Prog      Date        Description
' MO        03/07/02    ' BMIDS00054 Added new LV_PRODUCTS menu items and reorganised the sort
' CL        01/10/02    ' BMIDS00201 Added new LV_PRODUCTS menu SPECIAL_CONDITIONS
' DB        05/11/02    ' BMIDS00201 Re-ordered Promote and View Errors.
' DJP       22/02/03    ' BM0318 Added Third Party menu.
' DJP       06/03/03    ' BM0282 Add Locks menu items.

' MARS change history
'--------------------
' PJO       25/11/2005  MAR81 Add Activate to LV Addresses
' ~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~

' EPSOM Specific
' Prog      Date        Description
' pct       09/03/2006  EP15 Intermediaries = Packager or Broker - only one level.
' PB        17/10/2006  EP2_14  E2CR20 - Unit organisation changes
' TW        18/12/2006  EP2_568 Add functionality to select which Omiga Users are returned
' TW        15/01/2007  EP2_826 Rationalisation of pop-up menus and actions to improve consistency and usability
' TW        22/02/2007  EP2_1577 - Mortgage Product "Errors" pop-up not displayed
' ~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''


' Menu control ID's
Option Explicit

Public Const FILE_MENU = 0
Public Const FILE_NEW_LENDER = 0
Public Const FILE_NEW_PRODUCT = 1
Public Const FILE_NEW_STAMP_DUTY = 2
Public Const FILE_NEW_MIG_RATE = 3
Public Const FILE_NEW_ADMIN_FEE = 4
Public Const FILE_NEW_VALUATION_FEE = 5
Public Const FILE_NEW_BASE_RATE = 6
Public Const FILE_NEW_ADDRESS = 7
Public Const FILE_NEW_ADDRESS_PANEL = 0
Public Const FILE_NEW_ADDRESS_PANEL_LEGAL_REP = 0
Public Const FILE_NEW_ADDRESS_PANEL_VALUER = 1
Public Const FILE_NEW_ADDRESS_PANEL_LENDER = 2
Public Const FILE_NEW_ADDRESS_LOCAL = 1
Public Const FILE_NEW_SYS_ADMIN = 8
Public Const FILE_NEW_COMBOBOX = 9
Public Const FILE_NEW_ERROR = 10

Public Const EDIT_MENU = 1
Public Const EDIT_EDIT = 0
Public Const EDIT_COPY = 1
Public Const EDIT_DELETE = 2
'Public Const EDIT_VIEW = 1
'Public Const EDIT_COPY = 2

Public Const VIEW_MENU = 2
Public Const VIEW_REFRESH = 0

Public Const TOOLS_MENU = 3
Public Const TOOLS_FIND_LENDER = 0
Public Const TOOLS_FIND_PRODUCT = 1
Public Const TOOLS_FIND_STAMP_DUTY = 2
Public Const TOOLS_FIND_MIG_RATE = 3
Public Const TOOLS_FIND_ADMIN_FEE = 4
Public Const TOOLS_FIND_VALUATION_FEE = 5
Public Const TOOLS_FIND_BASE_RATE = 6
Public Const TOOLS_FIND_ADDRESS = 7
Public Const TOOLS_FIND_ADDRESS_PANEL = 0
Public Const TOOLS_FIND_ADDRESS_PANEL_LEGAL_REP = 0
Public Const TOOLS_FIND_ADDRESS_PANEL_VALUER = 1
Public Const TOOLS_FIND_ADDRESS_PANEL_LENDER = 2
Public Const TOOLS_FIND_ADDRESS_LOCAL = 1
Public Const TOOLS_FIND_SYS_ADMIN = 8
Public Const TOOLS_FIND_COMBOBOX = 9
Public Const TOOLS_FIND_ERROR = 10

Public Const HELP_MENU = 4
Public Const HELP_CONTENTS = 0
Public Const HELP_SEARCH = 1
Public Const HELP_ABOUT = 2

' Not continuous because of the separators.
Public Const LV_LENDERS_MENU = 5
Public Const LV_LENDERS_EDIT = 0
Public Const LV_LENDERS_PROMOTE = 1
Public Const LV_LENDERS_CONTACT_DETAILS = 3
Public Const LV_LENDERS_PARAMETERS = 4
Public Const LV_LENDERS_LEGAL_FEES = 5
Public Const LV_LENDERS_OTHER_FEES = 6
Public Const LV_LENDERS_MIG_RATES = 7
Public Const LV_LENDERS_LEDGER_CODES = 8

Public Const LV_PRODUCTS_MENU = 6

Public Const LV_PRODUCTS_EDIT = 0
Public Const LV_PRODUCTS_COPY = 1
Public Const LV_PRODUCTS_DELETE = 2

' TW 22/02/2007 EP2_1577
'Public Const LV_PRODUCTS_CHANNEL_ELIG = 4
'Public Const LV_PRODUCTS_EMPLOYMENT_ELIG = 5
'Public Const LV_PRODUCTS_PROPERTY_LOCATION = 6
'Public Const LV_PRODUCTS_PURPOSE_OF_LOAN = 7
'Public Const LV_PRODUCTS_TYPE_OF_APP_ELIG = 8
'Public Const LV_PRODUCTS_TYPE_OF_BUYER_ELIG = 9
'
'Public Const LV_PRODUCTS_ADMIN_FEE = 11
'Public Const LV_PRODUCTS_INCENTIVE_SETS = 12
'Public Const LV_PRODUCTS_REDEMPTION_FEES = 13
'Public Const LV_PRODUCTS_VALUATION_FEE = 14
'Public Const LV_PRODUCTS_OTHER_FEES = 15
'
'Public Const LV_PRODUCTS_INTEREST_RATES = 17
'Public Const LV_PRODUCTS_TYPICAL_APR = 18
'Public Const LV_PRODUCTS_SPECIAL_GROUPS = 19
'Public Const LV_PRODUCTS_MIG_RATES = 20
'Public Const LV_PRODUCTS_INCOME_MULTIPLES = 21
'Public Const LV_PRODUCTS_PARAMETERS_BANDS_SETS = 22
'Public Const LV_PRODUCTS_SPECIAL_CONDITIONS = 23
' TW 22/02/2007 EP2_1577 End

'Public Const LV_PRODUCTS_PROMOTE = 24
'Public Const LV_PRODUCTS_VIEW_ERRORS = 25

' TW 15/01/2007 EP2_826
'Public Const LV_PRODUCTS_PROMOTE = 25
'Public Const LV_PRODUCTS_VIEW_ERRORS = 26
Public Const LV_PRODUCTS_PROMOTE = 4
Public Const LV_PRODUCTS_VIEW_ERRORS = 5
' TW 15/01/2007 EP2_826 End



Public Const LV_PRODUCTS_BASE_RATE_SETS = 27

'Public Const LV_PRODUCTS_DELETE = 13

Public Const LV_EDIT_MENU = 7

Public Const LV_EDIT = 0
Public Const LV_DELETE = 2
Public Const LV_MARK_FOR_PROMOTION = 3
' TW 15/01/2007 EP2_826
Public Const LV_BASERATE_MENU = 18
Public Const LV_BASERATE_VIEW = 0
Public Const LV_BASERATE_DELETE = 2
Public Const LV_BASERATE_MARK_FOR_PROMOTION = 3
' TW 15/01/2007 EP2_826 End

Public Const TV_OPTIONS_MENU = 8
Public Const TV_OPTIONS_NEW = 0
Public Const TV_OPTIONS_FIND = 1
'PB 13/10/2006 Begin
Public Const TV_OPTIONS_RETRIEVE_ALL = 2
Public Const TV_OPTIONS_RETRIEVE_OMIGA_UNITS = 3
Public Const TV_OPTIONS_RETRIEVE_INTRO_UNITS = 4
' PB 13/10/2006 End
' TW 18/12/2006 EP2_568
Public Const TV_OPTIONS_RETRIEVE_OMIGA_USERS = 5
Public Const TV_OPTIONS_RETRIEVE_INTRO_USERS = 6
' TW 18/12/2006 EP2_568 End

Public Const LV_ADDRESS_MENU = 9
Public Const LV_ADDRESS_EDIT = 0
Public Const LV_ADDRESS_ACTIVATE = 1
Public Const LV_ADDRESS_CONTACT_DETAILS = 3
Public Const LV_ADDRESS_LEGAL_REP = 4
Public Const LV_ADDRESS_VALUER = 5
Public Const LV_ADDRESS_BANK_DETAILS = 6
Public Const LV_ADDRESS_DELETE = 8
Public Const LV_ADDRESS_PROMOTE = 9

Public Const LV_USER_MENU = 11
Public Const LV_USER_EDIT = 0
Public Const LV_USER_DELETE = 1
Public Const LV_USER_QUALIFICATIONS = 2
Public Const LV_USER_COMPETENCY_HISTORY = 3
Public Const LV_USER_USER_ROLES = 4
Public Const LV_USER_PROMOTE = 6

Public Const LV_LOCK_MENU = 12
Public Const LV_FIND_LOCK = 0
' DJP BM0282 Add menu item for searching for all locks.
Public Const LV_ALL_LOCK = 1
Public Const LV_CREATE_LOCK = 2
Public Const LV_REMOVE_LOCK = 3

Public Const LV_APPLICATION_PROCESSING = 13
Public Const LV_CANCEL_APPLICATION = 0

'AA 13/02/01 - Batch Scheduler

Public Const LV_BATCH_MENU = 14
Public Const LV_BATCH_PROCESS_VIEW = 0
Public Const LV_BATCH_PROCESS_DELETE = 1
Public Const LV_BATCH_PROCESS_CANCEL = 3
Public Const LV_BATCH_PROCESS_RESTART = 4
Public Const LV_BATCH_PROCESS_VIEWJOBS = 6
Public Const LV_BATCH_PROCESS_LAUNCH = 8

'AA 06/04/01 - Intermediaries
Public Const LV_INTERMEDIARY_MENU = 15
Public Const LV_INTERMEDIARY_MENU_ADD = 0
Public Const LV_INTERMEDIARY_MENU_REFRESH = 1
Public Const LV_INTERMEDIARY_MENU_FIND = 2
Public Const LV_INTERMEDIARY_ADMIN = 0
Public Const LV_INTERMEDIARY_COMPANY = 1
Public Const LV_INTERMEDIARY_INDIVIDUAL = 2
Public Const LV_INTERMEDIARY_LEADAGENT = 3
Public Const LV_INTERMEDIARY_PACKAGER = 4 ' EP15 pct
Public Const LV_INTERMEDIARY_BROKER = 5 ' EP15 pct

' DJP BM0318
Public Const LV_THIRD_PARTY_MENU = 16
Public Const LV_THIRD_PARTY_MENU_ADD = 0
Public Const LV_THIRD_PARTY_MENU_FIND = 1
Public Const LV_THIRD_PARTY_MENU_ALL = 2

'   BMids
'
' AW    23/05/02    BM017/087/088
Public Const FILE_NEW_REDEMTION_FEE = 15
Public Const FILE_NEW_MP_MIG_RATE = 16
Public Const FILE_NEW_INCOME_MULT = 17

' PB 17/10/2006
Public Enum LV_DETAIL
    LV_DETAIL_NONE = -1
    LV_DETAIL_ALL = 0
    LV_DETAIL_OMIGA_UNITS = 1
    LV_DETAIL_INTRO_UNITS = 2
End Enum
' TW 18/12/2006 EP2_568
Public Enum LV_USERS
    LV_USERS_NONE = -1
    LV_USERS_ALL = 0
    LV_USERS_OMIGA_USERS = 1
    LV_USERS_INTRO_USERS = 2
End Enum
' TW 18/12/2006 EP2_568 End


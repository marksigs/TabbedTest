Attribute VB_Name = "Constants"
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' Description   - Constants used throughout Supervisor
'
' Prog      Date        Description
' DJP       09/11/00    Phase 2 Task Management
' CL        24/04/01    Modification for inclusion of frmEditCurrencies
' DJP       26/06/01    SQL Server port
' DJP       20/11/01    SYS2831/SYS2912 Support client variants & SQL Server
'                       locking problem. Added key constants.
' STB       07/12/01    SYS2550 - Add Intermediaries to Key Constants.
' STB       21/12/01    SYS2550 - Removed un-used Intermediary globals.
' STB       21/01/02    SYS2957 Supervisor Security Enhancement.
' DJP       15/02/02    SYS4094 Added NoTab constant.
' STB       13/05/02    SYS4417 Added AllowableIncomeFactors.
' SA        10/06/02    SYS4497 New Batch Status - In Progress added.
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' BMIDS SPECIFIC
' Prog      Date        Description
' GD        06/06/2002  BMIDS00016 Promotions
' DB        06/01/2003  SYS5457 Modify the promotion of Intermediaries
' DJP       22/02/2003  BM0318 Left Mouse is 1, not 0
' IK        07/03/03    BM0314 Added remove (DMS) Document Locks
' JD        14/06/04    BMIDS765 CC076 added rental income rates
' JD        30/03/2005  BMIDS982 Changed screen text from MIG to HLC
' JD        06/04/2005  BMIDS982 Changed screen text from Redemption Fee to ERC
' ~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
' MARS Specific

' Prog      Date        Description
' GHun      16/08/2005  MAR45 Apply BBG1370 (New screen for print configuration)
' GHun      14/10/2005  MAR202 Added new Pack screen

' ~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
' EPSOM Specific

' Prog      Date        Description
' pct       09/03/2006  EP15 Intermediaries = Packager or Broker - only one level
' TW        09/10/2006  EP2_7 Added additiona; fee types
' PB        17/10/2006  EP2_14  E2CR20 - Unit organisation changes
' PB        20/10/2006  EP2_13  E2CR20 - as above
' TW        18/11/2006  EP2_132 ECR20/21 - Added new constants for promotions
' TW        23/11/2006  EP2_172 Change control EP2_5 - E2CR16 changes related to Introducer/Product Exclusives
' TW        30/11/2006  EP2_253 - Changes related to Mortgage Product Application Eligibility
' TW        05/12/2006  EP2_258/EP2_194 - Change to Principal Firms to read Principal Firms/Networks
' TW        11/12/2006  EP2_20 WP1 - Loans & Products Supervisor Changes part 3
' TW        14/12/2006  EP2_518 Procuration Fees Support
' TW        05/02/2007  EP2_706 - Should  network be mandatory data for ar firms
' TW        27/11/2007  DBM594 - Add new functionality "Payments for Completion"
' ~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
Option Explicit

Public Const BATCH_PROCESS = "Batch"
Public Const LEFT_MOUSE As Integer = 1
Public Const RIGHT_MOUSE As Integer = 2
Public Const MAX_SHORT As Long = 32767
Public Const LV_BATCH_PROCESS_MENU  As String = "HHHH"
Public Const SUPERVISOR_KEY As String = "Software\Omiga4Supervisor"
Public Const APP_SERVER As String = "ApplicationServer"
Public Const APP_SERVER_DIR As String = "ApplicationServerDirectory"
Public Const LV_BATCH_PROCESS_EDIT As String = "OK"
Public Const REG_SERVICE_NAME As String = "ServiceName"
Public Const REG_USER_ID As String = "UserID"
Public Const REG_PASSWORD As String = "Password"
Public Const REG_PROVIDER As String = "Provider"
Public Const REG_APP_SERVER As String = "AppServer"
Public Const REG_ACTIVE As String = "Active"
Public Const REG_CONNECTION_STRING As String = "ConnectionString"
Public Const REG_DATABASE_TYPE As String = "DatabaseType"
Public Const REG_DATABASE_SERVER As String = "DatabaseServer"
Public Const LENDERS As String = "Lenders"
Public Const PRODUCTS As String = "Products"
Public Const MORTGAGE_PRODUCTS As String = "Mortgage Products"
Public Const STAMP_DUTY_RATES As String = "Stamp Duty Rates"
Public Const MIG_RATE_SETS As String = "HLC Rate Sets"
' TW 09/10/2006 EP2_7
Public Const ADDITIONAL_BORROWING_FEES As String = "Additional Borrowing Fee Sets"
Public Const CREDIT_LIMIT_INCREASE_FEES As String = "Credit Limit Increase Fee Sets"
' TW 09/10/2006 EP2_7 End
' TW 11/12/2006 EP2_20
Public Const TRANSFER_OF_EQUITY_FEES As String = "Transfer of Equity Fee Sets"
' TW 11/12/2006 EP2_20 End
' TW 14/12/2006 EP2_518
Public Const PROCURATION_FEES As String = "Procuration Fees"
Public Const DEFAULT_PROCURATION_FEES As String = "Default Procuration Fees"
Public Const LOAN_AMOUNT_ADJUSTMENTS As String = "Loan Amount Adjustments"
Public Const LTV_AMOUNT_ADJUSTMENTS As String = "LTV Amount Adjustments"
' TW 14/12/2006 EP2_518 End


Public Const ADMIN_FEES As String = "Admin Fee Sets"
Public Const VALUATION_FEES As String = "Valuation Fee Sets"
Public Const BASE_RATES As String = "Base Rate Sets"
Public Const NAMES_AND_ADDRESSES As String = "Third Parties"
Public Const LOCAL_ADDRESS As String = "Other"
Public Const PANEL_ADDRESS As String = "Panel"
Public Const LEGAL_REP_ADDRESS As String = "Legal Rep."
Public Const VALUER_ADDRESS As String = "Valuers"
Public Const LENDER_ADDRESS As String = "Lender"
Public Const SYSTEM_PARAMETERS As String = "Global Parameters"
' TW 10/01/2007 EP2_637
'Public Const ERROR_MESSAGES As String = "Error Messages"
Public Const ERROR_MESSAGES As String = "Messages"
' TW 10/01/2007 EP2_637 End
Public Const COMBOBOX_ENTRIES As String = "Combo Groups"
Public Const RATES_AND_FEES As String = "Rate and Fee Sets"
Public Const GLOBAL_PARAM_FIXED As String = "Fixed Parameters"
Public Const GLOBAL_PARAM_BANDED As String = "Banded Parameters"
Public Const ORGANISATION As String = "Organisation"
Public Const DISTRIBUTION_CHANNELS As String = "Distribution Channels"
Public Const DEPARTMENTS As String = "Departments"
Public Const REGIONS As String = "Regions"
Public Const COMPETENCIES As String = "Competencies"
Public Const WORKING_HOURS As String = "Working Hours"
Public Const UNITS As String = "Units"
Public Const APPLICATION_PROCESSING As String = "Application Processing"
Public Const PROMOTIONS As String = "Promotions"
Public Const DB_CONNECTIONS As String = "Database Connections"
Public Const INCOME_FACTORS As String = "Income Factors"
Public Const PRINTING As String = "Printing"    'MAR45 GHun
Public Const TEMPLATE As String = "Template"    'MAR45 GHun
Public Const PACK As String = "Pack"            'MAR202 GHun

'*=[MC]BMIDS763 - CC075 - FEESETS
Public Const PRODUCT_SWITCH_FEESETS As String = "Product Switch Fee Set"
Public Const INSURANCE_ADMIN_FEESETS As String = "Insurance Admin Fee Sets"
Public Const TT_FEESETS As String = "TT Fee Sets"
'*=[MC] END BMIDS763

'BMIDS765 JD
Public Const RENTAL_INCOME_RATES As String = "Rental Income Interest Rate Sets"

' DJP Phase 2 Task Management
Public Const TASK_MANAGEMENT As String = "Task Management"
Public Const TASK_MANAGEMENT_TASKS As String = "Tasks"
Public Const TASK_MANAGEMENT_STAGES As String = "Stages"
Public Const TASK_MANAGEMENT_ACTIVITIES As String = "Activities"

Public Const USERS As String = "Users"
Public Const LIFE_COVER_RATES As String = "Life Cover Rates"
Public Const BUILDINGS_AND_CONTENTS_PRODUCTS As String = "Buildings and Contents Products"
Public Const PAYMENT_PROTECTION_RATES As String = "Payment Protection Rates"
Public Const PAYMENT_PROTECTION_PRODUCTS As String = "Payment Protection Products"
Public Const COUNTRIES As String = "Countries"

Public Const LV_OPTIONS As String = "Options"
Public Const LOCK_MAINTENANCE As String = "Lock Maintenance"
Public Const LOCKS_ONLINE As String = "Online"
Public Const LOCKS_OFFLINE As String = "Offline"

Public Const LOCKS_ONLINE_APPLICATION As String = "Online Application"
Public Const LOCKS_ONLINE_CUSTOMER As String = "Online Customer"

Public Const LOCKS_OFFLINE_APPLICATION As String = "Offline Application"
Public Const LOCKS_OFFLINE_CUSTOMER As String = "Offline Customer"

' ik_bm0314
Public Const LOCKS_DOCS As String = "Documents"

Public Const SUPERVISOR_INVALID_PASSWORD As Long = vbObjectError + 100

Public Const STAGE_CANCELLED As String = "C"
Public Const STAGE_COMPLETE As String = "MC"
Public Const APPSTAGE_COMBO As String = "APPLICATIONSTAGE"

Public Const DELETE_FLAG_SET As Integer = 1
Public Const DELETE_FLAG_NOT_SET As Integer = 0

Public Const RECURRING_NONWORKING_DAY As String = "R"

'AA 26/01/01
Public Const ADDITIONAL_QUESTIONS As String = "Questions"
Public Const CONDITIONS As String = "Conditions"

'AA 13/02/01 - Batch Scheduler
Public Const BATCH_SCHEDULER As String = "Batch Scheduler"
' TW 27/11/2007 DBM594
Public Const BATCH_PROCESSING As String = "Batch Processing"
Public Const PAYMENTS_FOR_COMPLETION As String = "Payments for Completion"
' TW 27/11/2007 DBM594 End

'AA 14/02/01 - BaseRate
'SA 10/06/02 - New status of In-Progress added
Public Const BASE_RATE As String = "Base Rate"
Public Const PAYMENT_PROCESSING_TYPE As String = "P"
Public Const BATCH_STATUS_CREATED As String = "CR"
Public Const BATCH_STATUS_LAUNCHED As String = "L"
Public Const BATCH_STATUS_CANCELLED As String = "CA"
Public Const BATCH_STATUS_COMPLETED As String = "CO"
Public Const BATCH_STATUS_INPROGRESS As String = "I"
Public Const BATCH_SCHEDULE_STATUS_COMPLETED As String = "C"
Public Const BATCH_SCHEDULE_STATUS_RUNNING As String = "R"
Public Const BATCH_SCHEDULE_STATUS_WAITING As String = "W"
Public Const BATCH_SCHEDULE_STATUS_CANCELLED As String = "CA"
Public Const PAY_SANCTIONED_PAYMENTS As String = "P"

'AA 05/03/01 - Printing Template

Public Const PRINTING_TEMPLATE As String = "Printing Template"
Public Const PRINTING_DOCUMENT As String = "Document Locations"   'MAR45 GHun
Public Const PRINTER_LOCATION_REMOTE As String = "R"
Public Const PRINTER_LOCATION_FILE As String = "F"
Public Const PRINTING_PACK As String = "Pack Definition"    'MAR202 GHun

'AA 09/03/01
Public Const CASE_TRACKING As String = "Case Tracking"
Public Const BUSINESS_GROUPS As String = "Business Groups"

'AA 06/04/01 - INTERMEDIARIES
Public Const INTERMEDIARIES As String = "Intermediaries"

Public Const LEADAGENT As String = "Lead Agent"
Public Const INDIVIDUAL As String = "Individual"
Public Const INTERMEDIARY_COMPANY As String = "Company"
Public Const ADMIN_CENTRE As String = "Administration Centre"
Public Const PACKAGER As String = "Packager" ' EP15 pct
Public Const BROKER As String = "Broker" ' EP15 pct

'PB 17/10/2006 EP2_13
Public Const INTRODUCERS As String = "Introducers"
Public Const INTRODUCER As String = "Introducer"

Public Const ASSOCIATIONS As String = "Associations"
Public Const CLUBS As String = "Clubs"
Public Const FIRMS As String = "Firms"
Public Const INDIVIDUALS As String = "Individuals"

'PB 03/11/2006 EP2_13 Redesign Begin
Public Const PRINCIPALS As String = "Principal Firms/Networks"
Public Const PACKAGERS As String = "Packagers"
Public Const ARFIRMS As String = "AR Firms"
Public Const ARBROKERS As String = "AR Brokers"
Public Const DABROKERS As String = "DA Brokers"
'PB 03/11/2006 EP2_13 Redesign End

' TW 18/11/2006 EP2_132
Public Const FIRMTRADINGNAME As String = "Firm Trading Name"
Public Const CLUBSANDASSOCIATIONS As String = "Clubs/Associations"
Public Const INTRODUCERFIRM As String = "Introducer Firm"
Public Const FIRMPERMISSIONS As String = "Firm Permissions"
Public Const NETWORKASSOCIATIONS As String = "Networks/Associations"
Public Const ACTIVITYFSA As String = "Activity FSA"
Public Const PRINCIPALFIRMPACKAGER As String = "Principal Firms/Packagers"
' TW 18/11/2006 EP2_132 End
' TW 23/11/2006 EP2_172
Public Const EXCLUSIVEPRODUCTS As String = "Exclusive Products"
' TW 23/11/2006 EP2_172 End
' TW 30/11/2006 EP2_253
Public Const MORTGAGEPRODUCTPRODUCTCLASS As String = "Mortgage Product Product Class"
Public Const MORTGAGEPRODUCTNATUREOFLOAN As String = "Mortgage Product Nature of Loan"
Public Const MORTGAGEPRODUCTINCOMESTATUS As String = "Mortgage Product Income Status"
' TW 30/11/2006 EP2_253 End
' TW 14/12/2006 EP2_518
Public Const DEFAULTPROCURATIONFEES As String = "Default Procuration Fees"
Public Const LOANAMOUNTADJUSTMENTS As String = "Loan Amount Adjustments"
Public Const LTVAMOUNTADJUSTMENTS As String = "LTV Amount Adjustments"
' TW 14/12/2006 EP2_518 End
' TW 05/02/2007 EP2_706
Public Const APPOINTMENTS As String = "Appointments"
' TW 05/02/2007 EP2_706 End

'PB EP2_13 End

'DB SYS5457 06/01/03 - These constants are used when storing the type of Intermediary for a promotion.
Public Const LEADAGENT_DESCRIPTION As String = INTERMEDIARIES & " - " & LEADAGENT
Public Const INDIVIDUAL_DESCRIPTION As String = INTERMEDIARIES & " - " & INDIVIDUAL
Public Const INTERMEDIARY_COMPANY_DESCRIPTION As String = INTERMEDIARIES & " - " & INTERMEDIARY_COMPANY
Public Const ADMIN_CENTRE_DESCRIPTION As String = INTERMEDIARIES & " - " & ADMIN_CENTRE
Public Const PACKAGER_DESCRIPTION As String = INTERMEDIARIES & " - " & PACKAGER ' EP15 pct
Public Const BROKER_DESCRIPTION As String = INTERMEDIARIES & " - " & BROKER ' EP15 pct

'AA Intermediary Validation Types
Public Const INTERMEDIARIES_ADMIN As String = "A"
Public Const INTERMEDIARIES_COMPANY As String = "C"
Public Const INTERMEDIARIES_INDIVIDUAL As String = "I"
Public Const INTERMEDIARIES_LEADAGENT As String = "L"
Public Const INTERMEDIARIES_PACKAGER As String = "P" ' EP15 pct
Public Const INTERMEDIARIES_BROKER As String = "B" ' EP15 pct

'CL Currency Calculator
Public Const CURRENCIES As String = "Currencies"
Public Const COMBO_NONE As String = "<Select>"

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' Key constants - Either positional indexes or literal keys to obtain key
' variables from the colKeys collections.
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Public Const ADDRESS_KEY            As String = "Address"
Public Const CONTACT_DETAILS_KEY    As String = "ContactDetails"
Public Const ORGANISATION_ID_KEY    As Long = 1
Public Const RATESET_KEY            As Long = 1
Public Const MIG_RATESET_KEY        As Long = 1
Public Const PRODUCT_CODE_KEY       As Long = 1
Public Const PRODUCT_START_DATE_KEY As Long = 2
Public Const INTERMEDIARY_KEY       As Long = 1
Public Const INCOMEFACTOR_ORGANISATIONID_KEY As Long = 1
Public Const INCOMEFACTOR_INCOMEGROUP_KEY    As Long = 2
Public Const INCOMEFACTOR_TYPE_KEY           As Long = 3

'These ObectID values correspond to securable items/resources which are added to
'the virtual table in the Security component. They are numbered to correspond to
'the order they appear in the tree structure.

'Promotions.
Public Const ID_PROMOTIONS              As Long = 100

'Combobox Entries.
Public Const ID_COMBOBOX_ENTRIES        As Long = 200

'Global Parameters.
Public Const ID_SYSTEM_PARAMETERS       As Long = 300
Public Const ID_GLOBAL_PARAM_FIXED      As Long = 310
Public Const ID_GLOBAL_PARAM_BANDED     As Long = 320

'Rates and Fees.
Public Const ID_RATES_AND_FEES          As Long = 400
' TW 09/10/2006 EP2_7
Public Const ID_ADDITIONAL_BORROWING_FEES    As Long = 405
Public Const ID_CREDIT_LIMIT_INCREASE_FEES   As Long = 407
' TW 09/10/2006 EP2_7 End
Public Const ID_ADMIN_FEES              As Long = 410
Public Const ID_VALUATION_FEES          As Long = 420
Public Const ID_BASE_RATES              As Long = 430
Public Const ID_BASE_RATE               As Long = 440
'*=[MC]BMIDS763 - CC075 FEESETS
Public Const ID_PRODUCT_SWITCH_FEESETS  As Long = 450
Public Const ID_INSURANCE_ADMIN_FEESETS As Long = 460
Public Const ID_TT_FEESETS              As Long = 470
' TW 11/12/2006 EP2_20
Public Const ID_TRANSFER_OF_EQUITY_FEES As Long = 475
' TW 11/12/2006 EP2_20 End
'*=[MC]SECTION END - BMIDS763
' BMIDS765 JD
Public Const ID_RENTALINCOMERATES       As Long = 480

'Batch Secheduler.
Public Const ID_BATCH_SCHEDULER         As Long = 500

'Lenders.
Public Const ID_LENDERS                 As Long = 600

'Products.
Public Const ID_PRODUCTS                    As Long = 700
Public Const ID_MORTGAGE_PRODUCTS           As Long = 710
Public Const ID_LIFE_COVER_RATES            As Long = 720
Public Const ID_BUILDINGS_AND_CONTENTS      As Long = 730
Public Const ID_PAYMENT_PROTECTION_RATES    As Long = 740
Public Const ID_PAYMENT_PROTECTION_PRODUCTS As Long = 750

'Organisation.
Public Const ID_ORGANISATION            As Long = 800
Public Const ID_COUNTRIES               As Long = 810
Public Const ID_DISTRIBUTION_CHANNELS   As Long = 820
Public Const ID_DEPARTMENTS             As Long = 830
Public Const ID_REGIONS                 As Long = 840
Public Const ID_UNITS                   As Long = 850
Public Const ID_COMPETENCIES            As Long = 860
Public Const ID_WORKING_HOURS           As Long = 870
Public Const ID_USERS                   As Long = 880

'Third Parties.
Public Const ID_NAMES_AND_ADDRESSES     As Long = 900
Public Const ID_PANEL_ADDRESS           As Long = 910
Public Const ID_LEGAL_REP_ADDRESS       As Long = 920
Public Const ID_VALUER_ADDRESS          As Long = 930
Public Const ID_LOCAL_ADDRESS           As Long = 940
'Income Factors.
Public Const ID_INCOME_FACTORS          As Long = 980

'Error Messages.
Public Const ID_ERROR_MESSAGES          As Long = 1000

'Locks.
Public Const ID_LOCK_MAINTENANCE            As Long = 1100
Public Const ID_LOCKS_ONLINE                As Long = 1110
Public Const ID_LOCKS_ONLINE_APPLICATION    As Long = 1120
Public Const ID_LOCKS_ONLINE_CUSTOMER       As Long = 1130
Public Const ID_LOCKS_OFFLINE               As Long = 1140
Public Const ID_LOCKS_OFFLINE_APPLICATION   As Long = 1150
Public Const ID_LOCKS_OFFLINE_CUSTOMER      As Long = 1160
' ik_bm0314
Public Const ID_LOCKS_DOCS                  As Long = 1170

'Application Processing.
Public Const ID_APPLICATION_PROCESSING  As Long = 1200

'Task Management.
Public Const ID_TASK_MANAGEMENT             As Long = 1300
Public Const ID_TASK_MANAGEMENT_TASKS       As Long = 1310
Public Const ID_TASK_MANAGEMENT_STAGES      As Long = 1320
Public Const ID_TASK_MANAGEMENT_ACTIVITIES  As Long = 1330
Public Const ID_CASE_TRACKING               As Long = 1330
Public Const ID_BUSINESS_GROUPS             As Long = 1340

'Questions.
Public Const ID_ADDITIONAL_QUESTIONS    As Long = 1400

'Conditions.
Public Const ID_CONDITIONS              As Long = 1500

'Printing Template.
'MAR45 GHun
'Public Const ID_PRINTING_TEMPLATE       As Long = 1600
'Printing
Public Const ID_PRINTING                As Long = 1600
Public Const ID_PRINTING_TEMPLATE       As Long = 1610
Public Const ID_PRINTING_DOCUMENT       As Long = 1620
Public Const ID_PRINTING_PACK           As Long = 1630  'MAR202 GHun
'MAR45 End

'Currencies.
Public Const ID_CURRENCIES              As Long = 1700

'Intermediaries.
Public Const ID_INTERMEDIARIES          As Long = 1800

'Introducers
Public Const ID_INTRODUCERS             As Long = 1900 'EP2_13 PB

' When no tab is selected
Public Const NoTab                      As Long = -1

'BMIDS
'   AW  15/05/02    BM088   Income Mulitiples
'   AW  16/05/02    BM087   MIG Rate Fee Sets
'   AW  21/05/02    BM017   Redemption Fee Sets


'Rates and Fees
Public Const INCOME_MULTIPLE As String = "Income Multiple Sets"
Public Const ID_INCOME_MULTIPLE             As Long = 450
'GD - for promotions 28/05/02 BMIDS00016 START
Public Const MP_MIG_RATE_SETS As String = "HLC Rate Sets"
Public Const MP_MIG_RATE_BAND As String = "HLC Rate Sets"
'GD - for promotions 28/05/02 BMIDS00016 END

Public Const ID_MP_MIG_RATE_SETS               As Long = 460
Public Const MP_MIG_RATESET_KEY        As Long = 1
'GD - for promotions 28/05/02 BMIDS00016 START
Public Const REDEM_FEE_SETS As String = "ERC Sets"  'JD BMIDS982
Public Const REDEM_FEE_BAND As String = "ERC Sets"  'JD BMIDS982
Public Const MORTGAGEPRODUCTCONDITION As String = "Mortgage Product Condition"
'GD - for promotions 28/05/02 BMIDS00016 END
Public Const ID_REDEM_FEE_SETS               As Long = 470
Public Const REDEM_FEE_SET_KEY        As Long = 1

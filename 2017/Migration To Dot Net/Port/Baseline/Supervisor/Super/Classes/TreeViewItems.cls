VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "TreeViewItems"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' Class         : TreeViewItems
' Description   : Provides functionality to add the TreeView to the Supervisor Main form.
' Change history
' Prog      Date        Description
' DJP       17/12/01    SYS2831 Client variants.
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Option Explicit

' Private data
Private m_colTreeData As New Collection
Private m_nodeCurrent As MSComctlLib.node
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' Function      :   SetupTreeValues
' Description   :   Sets up the treeview with all required nodes that need to
'                   be displayed. These are the Core treeview items that need to be displayed.
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Public Sub SetupTreeValues()
    On Error GoTo Failed
    Dim TreeItem As TreeAccess
    Dim nAddressIndex As Integer
    Dim nPanelIndex As Integer
    Dim nOptionsIndex As Integer
    Dim nRatesAndFeesIndex As Integer
    Dim nGlobalParamIndex As Integer
    Dim nLockIndex As Integer
    Dim nLockTypeIndex As Integer
    Dim nTaskManagementIndex As Integer
    Dim nCaseTrackingIndex As Integer
    Dim nIntermediariesIndex As Integer
    
    Dim mnuItem As Menu
    Set TreeItem = New TreeAccess
    
    frmMain.tvwDB.LabelEdit = 1
    frmMain.tvwDB.LineStyle = tvwRootLines
    frmMain.tvwDB.Sorted = False
    Set m_colTreeData = New Collection
    
    ' First the Root node
    Set TreeItem = New TreeAccess
    TreeItem.AddName LV_OPTIONS, , , 1
    m_colTreeData.Add TreeItem, LV_OPTIONS
    nOptionsIndex = m_colTreeData.Count()
    
    ' ComboBox Entries
    Set TreeItem = New TreeAccess
    TreeItem.AddName COMBOBOX_ENTRIES, tvwChild, nOptionsIndex
    TreeItem.AddHeading "Group Name", 30
    TreeItem.AddHeading "Notes", 50
    m_colTreeData.Add TreeItem, COMBOBOX_ENTRIES
    
    ' Global Parameters
    nGlobalParamIndex = m_colTreeData.Count + 1  ' + 1 because it will be added to the collection later, but for now we need the index
    Set TreeItem = New TreeAccess
    TreeItem.AddName SYSTEM_PARAMETERS, tvwChild, nOptionsIndex, nGlobalParamIndex
    TreeItem.AddHeading "Name", 15
    TreeItem.AddHeading "Data Type", 15
    TreeItem.AddHeading "Value", 15
    TreeItem.AddHeading "Active From", 15
    TreeItem.AddHeading "Active To", 15
    m_colTreeData.Add TreeItem, SYSTEM_PARAMETERS
    
    'Global Parameters -> Fixed
    Set TreeItem = New TreeAccess
    TreeItem.AddName GLOBAL_PARAM_FIXED, tvwChild, nGlobalParamIndex
    TreeItem.AddHeading "Name", 25
    TreeItem.AddHeading "Description", 40
    TreeItem.AddHeading "Start Date", 15
    TreeItem.AddHeading "Amount", 12
    TreeItem.AddHeading "Percentage", 13
    TreeItem.AddHeading "Maximum", 10
    TreeItem.AddHeading "Boolean", 10
    TreeItem.AddHeading "String", 15
    m_colTreeData.Add TreeItem, GLOBAL_PARAM_FIXED
    
    'Global Parameters -> Banded
    Set TreeItem = New TreeAccess
    TreeItem.AddName GLOBAL_PARAM_BANDED, tvwChild, nGlobalParamIndex
    TreeItem.AddHeading "Name", 20
    TreeItem.AddHeading "Description", 25
    TreeItem.AddHeading "Start Date", 15
    TreeItem.AddHeading "Band", 10
    TreeItem.AddHeading "Amount", 15
    TreeItem.AddHeading "Percentage", 15
    TreeItem.AddHeading "Maximum", 15
    TreeItem.AddHeading "Boolean", 15
    TreeItem.AddHeading "String", 15
    m_colTreeData.Add TreeItem, GLOBAL_PARAM_BANDED
    
    ' Rates and Fees
    Set TreeItem = New TreeAccess
    nRatesAndFeesIndex = m_colTreeData.Count() + 1
    TreeItem.AddName RATES_AND_FEES, tvwChild, nOptionsIndex, nRatesAndFeesIndex
    m_colTreeData.Add TreeItem, RATES_AND_FEES

    
    ' Rates and Fees -> Admin Fees
    Set TreeItem = New TreeAccess
    TreeItem.AddName ADMIN_FEES, tvwChild, nRatesAndFeesIndex
    TreeItem.AddHeading "Set Number", 20
    TreeItem.AddHeading "Start Date", 20
    TreeItem.AddHeading "Type of Application", 15
    TreeItem.AddHeading "Location", 15
    TreeItem.AddHeading "Fee Amount", 15
    m_colTreeData.Add TreeItem, ADMIN_FEES

    ' Rates and Fees -> Valuation Fees
    Set TreeItem = New TreeAccess
    TreeItem.AddName VALUATION_FEES, tvwChild, nRatesAndFeesIndex
    TreeItem.AddHeading "Set Number", 20
    TreeItem.AddHeading "Start Date", 20
    TreeItem.AddHeading "Type of Application", 15
    TreeItem.AddHeading "Location", 15
    TreeItem.AddHeading "Maximum Value", 15
    TreeItem.AddHeading "Valuation Fee", 15
    m_colTreeData.Add TreeItem, VALUATION_FEES

    ' Rates and Fees ->  BaseRate Sets
    Set TreeItem = New TreeAccess
    TreeItem.AddName BASE_RATES, tvwChild, nRatesAndFeesIndex
    TreeItem.AddHeading "Set Number", 25
    TreeItem.AddHeading "Start Date", 25
    TreeItem.AddHeading "Max Loan Amount", 20
    TreeItem.AddHeading "Max LTV", 20
    TreeItem.AddHeading "Base Rate", 15
    m_colTreeData.Add TreeItem, BASE_RATES
    
    ' Rates and Fees ->  Base Rates
    If g_clsMainSupport.DoesRateExist Then
        Set TreeItem = New TreeAccess
        TreeItem.AddName BASE_RATE, tvwChild, nRatesAndFeesIndex
        TreeItem.AddHeading "Rate ID", 15
        TreeItem.AddHeading "StartDate", 15
        TreeItem.AddHeading "Interest Rate", 15
        TreeItem.AddHeading "Type", 15
        TreeItem.AddHeading "Applied Date", 20
        m_colTreeData.Add TreeItem, BASE_RATE
    End If
    
    If g_clsMainSupport.DoesBatchExist Then
        'AA 13/01/02 - Batch Scheduler
        Set TreeItem = New TreeAccess
        TreeItem.AddName BATCH_SCHEDULER, tvwChild, nOptionsIndex
        TreeItem.AddHeading "Batch Number", 15
        TreeItem.AddHeading "Program Type", 20
        TreeItem.AddHeading "User ID", 15
        TreeItem.AddHeading "Batch Status", 20
        TreeItem.AddHeading "Batch Frequency", 15
        TreeItem.AddHeading "Batch Description", 30
        TreeItem.AddHeading "Execution Date/Time", 20
        m_colTreeData.Add TreeItem, BATCH_SCHEDULER
    End If
    
    ' Lenders
    Set TreeItem = New TreeAccess
    TreeItem.AddName LENDERS, tvwChild, nOptionsIndex
    TreeItem.AddHeading "Lender Code", 20
    TreeItem.AddHeading "Lender Name", 20
    TreeItem.AddHeading "Lender Start Date", 15
    TreeItem.AddHeading "Lender End Date", 15

    m_colTreeData.Add TreeItem, LENDERS

    ' Products
    Set TreeItem = New TreeAccess
    nAddressIndex = m_colTreeData.Count + 1
    TreeItem.AddName PRODUCTS, tvwChild, nOptionsIndex, nAddressIndex
    m_colTreeData.Add TreeItem, PRODUCTS

    ' Mortgage Products
    Set TreeItem = New TreeAccess
    TreeItem.AddName MORTGAGE_PRODUCTS, tvwChild, nAddressIndex
    TreeItem.AddHeading "Start Date", 10
    TreeItem.AddHeading "End Date", 10
    TreeItem.AddHeading "Product Code", 20
    TreeItem.AddHeading "Product Name", 30
    TreeItem.AddHeading "Lender Code", 20
    
    If Not g_clsVersion.DoesVersioningExist() Then
        TreeItem.AddHeading "Valid?", 20
    End If
    
    m_colTreeData.Add TreeItem, MORTGAGE_PRODUCTS

    ' Life Cover Rates
    Set TreeItem = New TreeAccess
    TreeItem.AddName LIFE_COVER_RATES, tvwChild, nAddressIndex
    TreeItem.AddHeading "Life Cover Number", 15
    TreeItem.AddHeading "Start Date", 15
    TreeItem.AddHeading "Cover Type", 20
    TreeItem.AddHeading "Applicant Gender", 10
    TreeItem.AddHeading "Max Age", 10
    TreeItem.AddHeading "Max Term", 10
    TreeItem.AddHeading "Annual Rate", 10
    TreeItem.AddHeading "Smoker Rate", 10
    TreeItem.AddHeading "Poor Health Rate", 10
    m_colTreeData.Add TreeItem, LIFE_COVER_RATES

    ' Buildings and Contents
    Set TreeItem = New TreeAccess
    TreeItem.AddName BUILDINGS_AND_CONTENTS_PRODUCTS, tvwChild, nAddressIndex
    TreeItem.AddHeading "Product Name", 25
    TreeItem.AddHeading "Product Number", 15
    TreeItem.AddHeading "Start Date", 15
    TreeItem.AddHeading "End Date", 15
    TreeItem.AddHeading "Valuables Limit", 10
    m_colTreeData.Add TreeItem, BUILDINGS_AND_CONTENTS_PRODUCTS

    ' Payment Protection
    Set TreeItem = New TreeAccess
    TreeItem.AddName PAYMENT_PROTECTION_RATES, tvwChild, nAddressIndex
    TreeItem.AddHeading "PP Rate Number", 15
    TreeItem.AddHeading "Start Date", 15
    TreeItem.AddHeading "Channel", 15
    TreeItem.AddHeading "Applicant Gender", 10
    TreeItem.AddHeading "Max Applicant Age", 15
    TreeItem.AddHeading "ASU Rate", 15
    TreeItem.AddHeading "U Rate", 15
    m_colTreeData.Add TreeItem, PAYMENT_PROTECTION_RATES

    ' Payment Protection Products
    Set TreeItem = New TreeAccess
    TreeItem.AddName PAYMENT_PROTECTION_PRODUCTS, tvwChild, nAddressIndex
    TreeItem.AddHeading "Product Code", 15
    TreeItem.AddHeading "Start Date", 15
    TreeItem.AddHeading "End Date", 15
    TreeItem.AddHeading "Product Name", 20
    TreeItem.AddHeading "Max Applicant Age", 15
    TreeItem.AddHeading "Rate Set", 15
    m_colTreeData.Add TreeItem, PAYMENT_PROTECTION_PRODUCTS
    
    ' Organisation
    Set TreeItem = New TreeAccess
    TreeItem.AddName ORGANISATION, tvwChild, nOptionsIndex
    m_colTreeData.Add TreeItem, ORGANISATION
    nAddressIndex = m_colTreeData.Count
    
    ' Country Entries
    Set TreeItem = New TreeAccess
    TreeItem.AddName COUNTRIES, tvwChild, nAddressIndex
    TreeItem.AddHeading "Country Number", 30
    TreeItem.AddHeading "Country Name", 50
    m_colTreeData.Add TreeItem, COUNTRIES
    
    ' Distribution Channels
    Set TreeItem = New TreeAccess
    TreeItem.AddName DISTRIBUTION_CHANNELS, tvwChild, nAddressIndex
    TreeItem.AddHeading "Channel Id", 25
    TreeItem.AddHeading "Channel Name", 75
    m_colTreeData.Add TreeItem, DISTRIBUTION_CHANNELS
    
    ' Departments
    Set TreeItem = New TreeAccess
    TreeItem.AddName DEPARTMENTS, tvwChild, nAddressIndex
    TreeItem.AddHeading "Department Id", 15
    TreeItem.AddHeading "Department Name", 30
    TreeItem.AddHeading "Channel", 15
    TreeItem.AddHeading "Active From", 10
    TreeItem.AddHeading "Active To", 10
    m_colTreeData.Add TreeItem, DEPARTMENTS
    
    ' Regions
    Set TreeItem = New TreeAccess
    TreeItem.AddName REGIONS, tvwChild, nAddressIndex
    TreeItem.AddHeading "Region Id", 15
    TreeItem.AddHeading "Region Type", 30
    TreeItem.AddHeading "Region Name", 30
    m_colTreeData.Add TreeItem, REGIONS
    
    ' Units
    Set TreeItem = New TreeAccess
    TreeItem.AddName UNITS, tvwChild, nAddressIndex
    TreeItem.AddHeading "Unit Id", 15
    TreeItem.AddHeading "Unit Name", 30
    TreeItem.AddHeading "Department Id", 30
    TreeItem.AddHeading "Active From", 10
    TreeItem.AddHeading "Active To", 10
    m_colTreeData.Add TreeItem, UNITS
    
    ' Competencies
    Set TreeItem = New TreeAccess
    TreeItem.AddName COMPETENCIES, tvwChild, nAddressIndex
    TreeItem.AddHeading "Competency Type", 15
    TreeItem.AddHeading "Active From", 10
    TreeItem.AddHeading "Active To", 10
    TreeItem.AddHeading "Funds Release Mandate", 15
    TreeItem.AddHeading "Loan Amount Mandate", 15
    TreeItem.AddHeading "LTV Mandate", 15
    TreeItem.AddHeading "Risk Assessment Mandate", 15
    m_colTreeData.Add TreeItem, COMPETENCIES
    
    ' Working Hours
    Set TreeItem = New TreeAccess
    TreeItem.AddName WORKING_HOURS, tvwChild, nAddressIndex
    TreeItem.AddHeading "Working Hours Type", 15
    TreeItem.AddHeading "Working Hours Type Description", 30
    TreeItem.AddHeading "Bank Holiday Indicator", 15
    m_colTreeData.Add TreeItem, WORKING_HOURS
    
    ' Users
    Set TreeItem = New TreeAccess
    TreeItem.AddName USERS, tvwChild, nAddressIndex
    TreeItem.AddHeading "User ID", 15
    TreeItem.AddHeading "User Name", 25
    TreeItem.AddHeading "Access Type", 20
    TreeItem.AddHeading "Active From", 15
    TreeItem.AddHeading "Active To", 15
    m_colTreeData.Add TreeItem, USERS
    
    ' Names and addresses
    Set TreeItem = New TreeAccess
    TreeItem.AddName NAMES_AND_ADDRESSES, tvwChild, nOptionsIndex
    m_colTreeData.Add TreeItem, NAMES_AND_ADDRESSES
    nAddressIndex = m_colTreeData.Count
    
    ' Names and addresses -> Panel
    Set TreeItem = New TreeAccess
    nPanelIndex = m_colTreeData.Count + 1
    TreeItem.AddName PANEL_ADDRESS, tvwChild, nAddressIndex, nPanelIndex
    m_colTreeData.Add TreeItem, PANEL_ADDRESS


    ' Names and addresses -> Legal Rep
    Set TreeItem = New TreeAccess
    TreeItem.AddName LEGAL_REP_ADDRESS, tvwChild, nPanelIndex
    TreeItem.AddHeading "Active From", 10
    TreeItem.AddHeading "Active To", 10
    TreeItem.AddHeading "Company Name", 25
'    treeItem.AddHeading "Organisation Type", 15
    TreeItem.AddHeading "Name and Address Type", 20
    TreeItem.AddHeading "Panel ID", 15
    m_colTreeData.Add TreeItem, LEGAL_REP_ADDRESS

    ' Names and addresses -> Valuer
    Set TreeItem = New TreeAccess
    TreeItem.AddName VALUER_ADDRESS, tvwChild, nPanelIndex
    TreeItem.AddHeading "Active From", 10
    TreeItem.AddHeading "Active To", 10
    TreeItem.AddHeading "Company Name", 25
'    treeItem.AddHeading "Organisation Type", 15
    TreeItem.AddHeading "Name and Address Type", 20
    TreeItem.AddHeading "Panel ID", 15
    m_colTreeData.Add TreeItem, VALUER_ADDRESS

    ' Names and addresses -> Local
    Set TreeItem = New TreeAccess
    TreeItem.AddName LOCAL_ADDRESS, tvwChild, nAddressIndex
    TreeItem.AddHeading "Active From", 10
    TreeItem.AddHeading "Active To", 10
    TreeItem.AddHeading "Company Name", 30
'    treeItem.AddHeading "Organisation Type", 20
    TreeItem.AddHeading "Name and Address Type", 25

    m_colTreeData.Add TreeItem, LOCAL_ADDRESS
    
    ' Error Messages
    Set TreeItem = New TreeAccess
    TreeItem.AddName ERROR_MESSAGES, tvwChild, nOptionsIndex
    TreeItem.AddHeading "Error Code ", 15
    TreeItem.AddHeading "Message Type", 15
    TreeItem.AddHeading "Message Text", 60
    m_colTreeData.Add TreeItem, ERROR_MESSAGES ' "Error Messages"
    
    ' Log Locks
    Set TreeItem = New TreeAccess
    nLockIndex = m_colTreeData.Count() + 1
    TreeItem.AddName LOCK_MAINTENANCE, tvwChild, nOptionsIndex, nLockIndex
    m_colTreeData.Add TreeItem, LOCK_MAINTENANCE


    ' Log Locks Online
    Set TreeItem = New TreeAccess
    nLockTypeIndex = m_colTreeData.Count() + 1
    TreeItem.AddName LOCKS_ONLINE, tvwChild, nLockIndex, nLockTypeIndex
    m_colTreeData.Add TreeItem, LOCKS_ONLINE


    ' Log Locks Online -> Application
    Set TreeItem = New TreeAccess
    TreeItem.AddName LOCKS_ONLINE_APPLICATION, tvwChild, nLockTypeIndex
    TreeItem.AddHeading "Application Number", 20
    TreeItem.AddHeading "Lock Date", 15
    'treeItem.AddHeading "Customer Number", 20
    TreeItem.AddHeading "User ID", 10
    TreeItem.AddHeading "Unit ID", 10
    TreeItem.AddHeading "Machine ID", 10
    m_colTreeData.Add TreeItem, LOCKS_ONLINE_APPLICATION

    ' Log Locks Online -> Customer
    Set TreeItem = New TreeAccess
    TreeItem.AddName LOCKS_ONLINE_CUSTOMER, tvwChild, nLockTypeIndex
    TreeItem.AddHeading "Customer Number", 20
    TreeItem.AddHeading "Lock Date", 15
    TreeItem.AddHeading "User ID", 10
    TreeItem.AddHeading "Unit ID", 10
    TreeItem.AddHeading "Machine ID", 10
    m_colTreeData.Add TreeItem, LOCKS_ONLINE_CUSTOMER

    ' Log Locks Offline
    Set TreeItem = New TreeAccess
    TreeItem.AddName LOCKS_OFFLINE, tvwChild, nLockIndex
    m_colTreeData.Add TreeItem, LOCKS_OFFLINE
    nLockTypeIndex = m_colTreeData.Count()

    ' Log Locks Offline -> Application
    Set TreeItem = New TreeAccess
    TreeItem.AddName LOCKS_OFFLINE_APPLICATION, tvwChild, nLockTypeIndex
    TreeItem.AddHeading "Application Number", 20
    TreeItem.AddHeading "Lock Date", 15
    TreeItem.AddHeading "User ID", 10
    TreeItem.AddHeading "Unit ID", 10
    TreeItem.AddHeading "Machine ID", 10
    m_colTreeData.Add TreeItem, LOCKS_OFFLINE_APPLICATION

    ' Log Locks Online -> Customer
    Set TreeItem = New TreeAccess
    TreeItem.AddName LOCKS_OFFLINE_CUSTOMER, tvwChild, nLockTypeIndex
    TreeItem.AddHeading "Customer Number", 20
    TreeItem.AddHeading "Lock Date", 15
    TreeItem.AddHeading "User ID", 10
    TreeItem.AddHeading "Unit ID", 10
    TreeItem.AddHeading "Machine ID", 10
    m_colTreeData.Add TreeItem, LOCKS_OFFLINE_CUSTOMER

    ' Application processing
    Set TreeItem = New TreeAccess
    TreeItem.AddName APPLICATION_PROCESSING, tvwChild, nOptionsIndex
    TreeItem.AddHeading "Application Number", 20
    TreeItem.AddHeading "Surname", 20
    TreeItem.AddHeading "ForeName", 15
    TreeItem.AddHeading "Address", 45
    
    If g_clsMainSupport.DoesTaskManagementExist() Then

        ' --- DJP Omiga 4 Phase 2 Task Management Start ---
        Set TreeItem = New TreeAccess
        nTaskManagementIndex = m_colTreeData.Count() + 1
        TreeItem.AddName TASK_MANAGEMENT, tvwChild, nOptionsIndex, nTaskManagementIndex
        m_colTreeData.Add TreeItem, TASK_MANAGEMENT

    
        ' Task Management -> Tasks
        Set TreeItem = New TreeAccess
        TreeItem.AddName TASK_MANAGEMENT_TASKS, tvwChild, nTaskManagementIndex
        TreeItem.AddHeading "Task ID", 15
        TreeItem.AddHeading "Task", 40
        TreeItem.AddHeading "Type", 20
        TreeItem.AddHeading "Owner Type", 20
        m_colTreeData.Add TreeItem, TASK_MANAGEMENT_TASKS
        
        ' Task Management -> Stages
        Set TreeItem = New TreeAccess
        TreeItem.AddName TASK_MANAGEMENT_STAGES, tvwChild, nTaskManagementIndex
        TreeItem.AddHeading "Stage ID", 15
        TreeItem.AddHeading "Stage", 40
        TreeItem.AddHeading "Authority Level", 20
        TreeItem.AddHeading "Exception?", 20
        m_colTreeData.Add TreeItem, TASK_MANAGEMENT_STAGES
        
        ' Task Management -> Activities
        Set TreeItem = New TreeAccess
        TreeItem.AddName TASK_MANAGEMENT_ACTIVITIES, tvwChild, nTaskManagementIndex
        TreeItem.AddHeading "Activity ID", 15
        TreeItem.AddHeading "Activity", 40
        TreeItem.AddHeading "Desciption", 40
        m_colTreeData.Add TreeItem, TASK_MANAGEMENT_ACTIVITIES
        ' --- DJP Omiga 4 Phase 2 Task Management End ---
        
        If g_clsMainSupport.DoesBusinessGroupsExist Then
            
            'Task Management -> Case Tracking
            Set TreeItem = New TreeAccess
            nCaseTrackingIndex = m_colTreeData.Count() + 1
            TreeItem.AddName CASE_TRACKING, tvwChild, nTaskManagementIndex, nCaseTrackingIndex
            m_colTreeData.Add TreeItem, CASE_TRACKING

            
            ' Case Tracking -> Business Groups
            Set TreeItem = New TreeAccess
            TreeItem.AddName BUSINESS_GROUPS, tvwChild, nCaseTrackingIndex
            TreeItem.AddHeading "Group Name", 25
            TreeItem.AddHeading "Business Area", 15
            m_colTreeData.Add TreeItem, BUSINESS_GROUPS
        End If
    Else
        m_colTreeData.Add TreeItem, APPLICATION_PROCESSING
    End If
    
    'AA - Questions
    Set TreeItem = New TreeAccess
    TreeItem.AddName ADDITIONAL_QUESTIONS, tvwChild, nOptionsIndex
    TreeItem.AddHeading "Reference", 30
    TreeItem.AddHeading "Question", 50
    TreeItem.AddHeading "Type", 30
    TreeItem.AddHeading "Details Required", 50
    TreeItem.AddHeading "Deleted", 30
    m_colTreeData.Add TreeItem, ADDITIONAL_QUESTIONS
    
    'AA - Conditions
    Set TreeItem = New TreeAccess
    TreeItem.AddName CONDITIONS, tvwChild, nOptionsIndex
    TreeItem.AddHeading "Reference", 30
    TreeItem.AddHeading "Name", 50
    TreeItem.AddHeading "Description", 50
    TreeItem.AddHeading "Type", 30
    TreeItem.AddHeading "Editable", 20
    TreeItem.AddHeading "Free Format", 20
    TreeItem.AddHeading "Rule Ref", 20
    TreeItem.AddHeading "Deleted", 20
    m_colTreeData.Add TreeItem, CONDITIONS
    
    'AA - Printing Template
    If g_clsMainSupport.DoesPrintTemplateExist Then
        Set TreeItem = New TreeAccess
        TreeItem.AddName PRINTING_TEMPLATE, tvwChild, nOptionsIndex
        TreeItem.AddHeading "TemplateId", 20
        TreeItem.AddHeading "DPS TemplateId", 20
        TreeItem.AddHeading "Document Group", 20
        TreeItem.AddHeading "Template Name", 20
        TreeItem.AddHeading "Description", 30
        TreeItem.AddHeading "Minumum Role", 20
        m_colTreeData.Add TreeItem, PRINTING_TEMPLATE
    End If
    
    If g_clsMainSupport.DoesCurrencyExist Then
        'CL - Currency Calculator
        Set TreeItem = New TreeAccess
        TreeItem.AddName CURRENCIES, tvwChild, nOptionsIndex
        TreeItem.AddHeading "Currency Code", 10
        TreeItem.AddHeading "Currency Name", 10
        TreeItem.AddHeading "Conversion Rate", 10
        TreeItem.AddHeading "Precision Rounding", 10
        TreeItem.AddHeading "Precision Direction", 10
        TreeItem.AddHeading "Base Currency", 10
        m_colTreeData.Add TreeItem, CURRENCIES
    End If
    
    'AA - Intermediary
    If g_clsMainSupport.DoesIntermediaryExist Then

        Set TreeItem = New TreeAccess
        nIntermediariesIndex = m_colTreeData.Count() + 1
        TreeItem.AddName INTERMEDIARIES, tvwChild, nOptionsIndex, nIntermediariesIndex
        TreeItem.AddHeading "Name", 20
        m_colTreeData.Add TreeItem, INTERMEDIARIES



        Set TreeItem = New TreeAccess
        TreeItem.AddName LEADAGENT, tvwChild, nIntermediariesIndex
        TreeItem.AddHeading "Company Name", 20
        TreeItem.AddHeading "Comission Number", 20
        TreeItem.AddHeading "Active From", 15
        TreeItem.AddHeading "Active To", 15
        TreeItem.AddHeading "Panel ID", 15
        m_colTreeData.Add TreeItem, LEADAGENT

        Set TreeItem = New TreeAccess
        TreeItem.AddName INDIVIDUAL, tvwChild, nIntermediariesIndex
        TreeItem.AddHeading "Title", 10
        TreeItem.AddHeading "Forename", 20
        TreeItem.AddHeading "Surname", 20
        TreeItem.AddHeading "Active From", 15
        TreeItem.AddHeading "Active To", 15
        TreeItem.AddHeading "Panel ID", 15
        m_colTreeData.Add TreeItem, INDIVIDUAL

    End If
    
    Exit Sub
Failed:
    g_clsErrorHandling.RaiseError Err.Number, Err.DESCRIPTION
End Sub
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' Function      :   AddTreeNodes
' Description   :   Given the collection m_colTreeData that contains all TreeView nodes to be added,
'                   loop through them all adding each item to the TreeView.
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Public Sub AddTreeNodes()
    On Error GoTo Failed
    Dim TreeItem As TreeAccess
    Dim nCount As Long
    Dim nRelationship As Integer
    Dim nParent As Integer
    Dim sName As String
    
    frmMain.tvwDB.Nodes.Clear
    Set m_nodeCurrent = Nothing
    
    For nCount = 1 To m_colTreeData.Count
        Set TreeItem = m_colTreeData.Item(nCount)
        nParent = TreeItem.GetParent()
        nRelationship = TreeItem.GetRelationship()
        sName = TreeItem.GetName()
        If (nRelationship = -1) Then
            AddNode sName
        Else
            AddNode sName, nRelationship, nParent
        End If
    Next nCount
    
    
    ' Expand the first node
    frmMain.tvwDB.Nodes(1).Expanded = True
    
    Exit Sub
Failed:
    g_clsErrorHandling.RaiseError Err.Number, Err.DESCRIPTION
End Sub
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' Function      :   AddNode
' Description   :   Given the collection m_colTreeData that contains all TreeView nodes to be added,
'                   loop through them all adding each item to the TreeView.
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub AddNode(ByRef sTitle As String, Optional vRelationship As Variant, Optional vParent As Variant)
    On Error GoTo Failed
    Dim mNode As node
    Dim nCount As Integer
    
    nCount = frmMain.tvwDB.Nodes.Count
    
    If (IsMissing(vRelationship) Or IsMissing(vParent)) Then
        Set mNode = frmMain.tvwDB.Nodes.Add(, , sTitle, sTitle, "closed")

    Else
        Set mNode = frmMain.tvwDB.Nodes.Add(vParent, vRelationship, sTitle, sTitle, "closed")
    End If
    
    Exit Sub
Failed:
    g_clsErrorHandling.RaiseError Err.Number, Err.DESCRIPTION
End Sub
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' Function      :   GetSelectedTreeKey
' Description   :   Returns the Key for the selected TreeView node.
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Public Function GetSelectedTreeKey() As Variant
    On Error GoTo Failed
    Dim mNode As node
    Dim clsTreeItem As TreeItem
    ' Which node was selected?
       
    Set mNode = frmMain.tvwDB.SelectedItem

    If (Not mNode Is Nothing) Then
        If TypeOf mNode.Tag Is TreeItem Then
            Set clsTreeItem = mNode.Tag
            GetSelectedTreeKey = clsTreeItem.GetType
        Else
            GetSelectedTreeKey = mNode.Key
        End If
    Else
        GetSelectedTreeKey = Null
    End If
    
    Exit Function
Failed:
    g_clsErrorHandling.RaiseError Err.Number, Err.DESCRIPTION
End Function
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' Function      :   GetTreeItem
' Description   :   Returns the item specified by sTag.
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Public Function GetTreeItem(sTag As String) As TreeAccess
    On Error GoTo Failed
    
    Set GetTreeItem = m_colTreeData(sTag)
    
    Exit Function
Failed:
    g_clsErrorHandling.RaiseError Err.Number, Err.DESCRIPTION
End Function
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' Function      :   SetTreeStatus
' Description   :   Sets the correct image for the node passed in, or the current node if not
'                   specified.
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Public Sub SetTreeStatus(Optional node As MSComctlLib.node = Nothing)
    On Error GoTo Failed
    If Not m_nodeCurrent Is Nothing Then
        m_nodeCurrent.Image = "closed"
    End If
    
    If node Is Nothing Then
        Set node = frmMain.tvwDB.SelectedItem
    End If
    
    If Not node Is Nothing Then
        If node.Expanded = False And node.children = False Then
            node.Image = "open"
            Set m_nodeCurrent = node
        ElseIf node.children = True Then
            node.Expanded = True
            node.Image = "open"
        End If
    End If
    
    Exit Sub
Failed:
    g_clsErrorHandling.RaiseError Err.Number, Err.DESCRIPTION
End Sub
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' Function      :   ClearCurrentTreeitem
' Description   :   Clears the current node.
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Public Sub ClearCurrentTreeitem()
    On Error GoTo Failed
    
    Set m_nodeCurrent = Nothing
    
    Exit Sub
Failed:
    g_clsErrorHandling.RaiseError Err.Number, Err.DESCRIPTION
End Sub
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' Function      :   GetTreeData
' Description   :   Returns the TreeView collection to be used for client specific changes i.e.,
'                   a client class could add their own TreeView items to the collection to extend
'                   the Core functionality.
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Public Function GetTreeData() As Collection
    On Error GoTo Failed
    
    If m_colTreeData Is Nothing Then
        g_clsErrorHandling.RaiseError errGeneralError, "Tree collection is empty"
    End If
    
    Set GetTreeData = m_colTreeData
    Exit Function
Failed:
    g_clsErrorHandling.RaiseError Err.Number, TypeName(Me) & ".GetTreeData " & Err.DESCRIPTION
End Function

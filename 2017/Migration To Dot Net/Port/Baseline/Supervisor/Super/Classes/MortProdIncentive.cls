VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "MortProdIncentive"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' Class         : MortProdIncentive
' Description   : Provides functionality for the Incentives tab on Mortgage Product.
' Change history
' Prog      Date        Description
' DJP       22/06/01    SYS2381 SQL Server port, and tidy up
' DJP       03/12/01    SYS2912 SQL Server locking problem.
' DJP       10/12/01    SYS2831 Client variants.
' DJP       07/01/02    SYS2831 Client variants.
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''

Option Explicit

' Constants
Private Const SET_KEY As Integer = 1
Private Const DATE_KEY As Integer = 2
Private Const INCENTIVE_TYPE As String = "Incentive Type"
Private Const INCENTIVE_BENEFIT_TYPE As String = "Incentive Benefit Type"

' Private data
Private m_bIsEdit As Boolean
Private m_sSetKey  As String
Private m_colProductKeys As Collection
Private m_colIncentiveGUIDs As Collection
Private m_sProductCode As String
Private m_sProductStartDate As String
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' Function      : Initialise
' Description   : Called from the Mortgage Product form to do initialisation of the Mortgage
'                 Product Incentives tab. Needs to read the Incentives from the database.
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Public Sub Initialise(bIsEdit As Boolean)
    On Error GoTo Failed
    Dim clsTableAccess As TableAccess
    
    Set m_colIncentiveGUIDs = New Collection
    m_bIsEdit = bIsEdit
    
    InitialiseFields
    SetupListViewHeaders
    
    If m_bIsEdit Then
        PopulateIncentives
    End If
    
    frmProductDetails.lvIncentives.LoadColumnDetails TypeName(Me)
    
    Exit Sub
Failed:
    g_clsErrorHandling.RaiseError Err.Number, Err.DESCRIPTION
End Sub
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' Function      : PopulateIncentives
' Description   : Popultates the Incentives ListView. Need to populate both Exclusive and Inclusive
'                 Incentives, and also Incentives that have been created but don't yet have a Mortgage
'                 Product Code and Start Date.
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub PopulateIncentives()
    On Error GoTo Failed
    Dim clsIncentives As MortProdIncentiveTable

    Set clsIncentives = New MortProdIncentiveTable

    ' Get Inclusive incentives
    clsIncentives.GetIncentives Inclusive, m_colProductKeys
    g_clsFormProcessing.PopulateFromRecordset frmProductDetails.lvIncentives, clsIncentives
    
    ' Get Exclusive incentives
    clsIncentives.GetIncentives Exclusive, m_colProductKeys
    g_clsFormProcessing.PopulateFromRecordset frmProductDetails.lvIncentives, clsIncentives, False

    ' Now get the incentives that have been created but don't yet have a product code or start date
    If Not m_colIncentiveGUIDs Is Nothing Then
        If m_colIncentiveGUIDs.Count > 0 Then
            ' Get Inclusive incentives
            clsIncentives.GetIncentives Inclusive, m_colProductKeys, , m_colIncentiveGUIDs
            g_clsFormProcessing.PopulateFromRecordset frmProductDetails.lvIncentives, clsIncentives, False
            
            ' Get Exclusive incentives
            clsIncentives.GetIncentives Exclusive, m_colProductKeys, , m_colIncentiveGUIDs
            g_clsFormProcessing.PopulateFromRecordset frmProductDetails.lvIncentives, clsIncentives, False
        End If
    End If
    
    Set frmProductDetails.lvIncentives.SelectedItem = Nothing

    Exit Sub
Failed:
    g_clsErrorHandling.RaiseError Err.Number, Err.DESCRIPTION
End Sub
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' Function      : SetupListViewHeaders
' Description   : Sets the column names and widths for the Incentives listview.
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub SetupListViewHeaders()
    On Error GoTo Failed
    Dim headers As New Collection
    Dim lvHeaders As listViewAccess

    lvHeaders.nWidth = 20
    lvHeaders.sName = INCENTIVE_TYPE
    headers.Add lvHeaders

    lvHeaders.nWidth = 40
    lvHeaders.sName = INCENTIVE_BENEFIT_TYPE
    headers.Add lvHeaders

    lvHeaders.nWidth = 20
    lvHeaders.sName = "Amount"
    headers.Add lvHeaders

    lvHeaders.nWidth = 15
    lvHeaders.sName = "Percentage"
    headers.Add lvHeaders

    lvHeaders.nWidth = 20
    lvHeaders.sName = "Percentage Max"
    headers.Add lvHeaders

    lvHeaders.nWidth = 30
    lvHeaders.sName = "Description"
    headers.Add lvHeaders

    frmProductDetails.lvIncentives.AddHeadings headers

    Exit Sub
Failed:
    g_clsErrorHandling.RaiseError Err.Number, Err.DESCRIPTION
End Sub
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' Function      : SetScreenFields
' Description   : Populates all fields on the Mortgage Product Incentives tab - called externaly but
'                 doesn't do anything for this tab.
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Public Function SetScreenFields()
    Dim bRet As Boolean
    On Error GoTo Failed
    bRet = True

    SetScreenFields = bRet
    
    Exit Function
Failed:
    g_clsErrorHandling.RaiseError Err.Number, Err.DESCRIPTION
End Function
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' Function      : SaveScreenData
' Description   : Saves all data on the Incentives tab
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Public Sub SaveScreenData()

    SetIncentiveProductDetails

    Exit Sub
Failed:
    g_clsErrorHandling.RaiseError Err.Number, Err.DESCRIPTION
End Sub
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' Function      : SetIncentiveProductDetails
' Description   : Saves the Product Code and Start Date to the Incentives we have created. Calls
'                 itself, once for Inclusive and once for Exclusive.
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub SetIncentiveProductDetails(Optional nType As IncentiveType = InclusiveAndExclusive)
    On Error GoTo Failed
    Dim clsIncentives As MortProdIncIncentiveTable
    
    Set clsIncentives = New MortProdIncIncentiveTable
    If nType = InclusiveAndExclusive Then
        SetIncentiveProductDetails Inclusive
        SetIncentiveProductDetails Exclusive
    Else
        If Not m_colIncentiveGUIDs Is Nothing Then
            If m_colIncentiveGUIDs.Count > 0 Then
                Dim clsBandedTable As BandedTable

                
                Set clsIncentives = New MortProdIncIncentiveTable
                clsIncentives.GetIncentives nType, m_colIncentiveGUIDs
                
                Set clsBandedTable = clsIncentives
                
                ' Now update with the Mortgage Product Code and Start Date
                clsBandedTable.SetUpdateValues m_colProductKeys
                clsBandedTable.SetUpdateSets
                clsBandedTable.DoUpdateSets
                TableAccess(clsIncentives).Update
                TableAccess(clsIncentives).CloseRecordSet

            End If
        End If
    End If
    
    Exit Sub
Failed:
    g_clsErrorHandling.RaiseError Err.Number, Err.DESCRIPTION
End Sub
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' Function      : Edit
' Description   : Called when the user presses the Edit button on the Product Details form,
'                 Incentives tab. Edit the Incentive - load the Incentives form.
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Public Sub Edit()
    On Error GoTo Failed
    Dim sType As String
    Dim colMatchValues As New Collection
    Dim clsTableAccess As TableAccess

    GetKeyMatchValues colMatchValues
    sType = GetType
    
    frmEditIncentives.SetIncentive colMatchValues, sType
    frmEditIncentives.SetIsEdit
    frmEditIncentives.Show vbModal, frmProductDetails

    If frmEditIncentives.GetReturnCode = MSGSuccess Then
        PopulateIncentives
        SetButtonState
    End If
    
    Unload frmEditIncentives
    
    Exit Sub
Failed:
    g_clsErrorHandling.RaiseError Err.Number, Err.DESCRIPTION
End Sub
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' Function      : Add
' Description   : Called when the user presses the Add button on the Product Details form,
'                 Incentives tab. Show the Incentives form.
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Public Sub Add()
    On Error GoTo Failed
    Dim clsIncentive As MortProdIncentiveTable
    
    frmEditIncentives.SetIsEdit False
    frmEditIncentives.Show vbModal, frmProductDetails

    If frmEditIncentives.GetReturnCode = MSGSuccess Then
        Set clsIncentive = frmEditIncentives.GetIncentives
        SaveNewIncentiveGUIDs clsIncentive
        
        PopulateIncentives
        SetButtonState
    End If
    
    Unload frmEditIncentives
    
    Exit Sub
Failed:
    g_clsErrorHandling.DisplayError
End Sub
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' Function      : Delete
' Description   : Called when the user presses the Delete button on the Product Details form,
'                 Incentives tab. Delete the Incentive - load the Incentives form.
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Public Sub Delete()
    On Error GoTo Failed
    Dim nType As IncentiveType
    Dim nResponse As Integer
    Dim colMatchValues As New Collection
    Dim clsTableAccess As TableAccess
    Dim clsIncentive As New MortProdIncentiveTable
    Dim clsIncentiveType As New MortProdIncIncentiveTable
    
    If Not frmProductDetails.lvIncentives.SelectedItem Is Nothing Then
        nResponse = MsgBox("Delete the selected record?", vbQuestion + vbYesNo)
        
        If nResponse = vbYes Then
            GetKeyMatchValues colMatchValues
            Set clsTableAccess = clsIncentiveType
            
            GetTypeIndex nType
            clsIncentiveType.SetType (nType)
            clsTableAccess.DeleteRow colMatchValues
            
            Set clsTableAccess = clsIncentive
            clsTableAccess.DeleteRow colMatchValues
    
            Set clsTableAccess = clsIncentiveType
            clsTableAccess.Update
            
            Set clsTableAccess = clsIncentive
            clsTableAccess.Update
    
            If m_colProductKeys Is Nothing Then
                Dim sKey As String
                Dim vIncentiveGUID As Variant
                
                vIncentiveGUID = colMatchValues(1)
                sKey = g_clsSQLAssistSP.ByteArrayToGuidString(CStr(vIncentiveGUID))
                
                m_colIncentiveGUIDs.Remove sKey
                
            End If
    
            If clsTableAccess.GetUpdated() = True Then
                PopulateIncentives
                SetButtonState
            End If
        End If
    End If
    
    Exit Sub
Failed:
    g_clsErrorHandling.RaiseError Err.Number, Err.DESCRIPTION
End Sub
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' Function      : SaveNewIncentiveGUIDs
' Description   : When the user has created new incentives when creating a new product (and hence
'                 a Product Code and Start Date don't exist yet) we need to save the GUIDs for future
'                 so we can update the Product Code and Start Date on OK. The GUIDs are saved in
'                 m_colIncentiveGUIDs.
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub SaveNewIncentiveGUIDs(clsIncentive As MortProdIncentiveTable)
    On Error GoTo Failed
    Dim rs As ADODB.Recordset
    Dim sKey As String
    Dim vIncentiveGUID As Variant
    
    If TableAccess(clsIncentive).RecordCount > 0 Then
        Set rs = TableAccess(clsIncentive).GetRecordSet
        
        rs.MoveFirst
        
        Do While Not rs.EOF
            vIncentiveGUID = clsIncentive.GetIncentiveGUID
            sKey = g_clsSQLAssistSP.ByteArrayToGuidString(CStr(vIncentiveGUID))
            m_colIncentiveGUIDs.Add vIncentiveGUID, sKey
            rs.MoveNext
        Loop
        TableAccess(clsIncentive).CloseRecordSet
    
    End If

    Exit Sub
Failed:
    g_clsErrorHandling.RaiseError Err.Number, Err.DESCRIPTION
End Sub
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' Function      : GetType
' Description   : Returns the Type of Incentive currently selected ("Inclusive" or "Exclusive")
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Function GetType() As String
    On Error GoTo Failed
    Dim sType As String
    sType = frmProductDetails.lvIncentives.getValueFromName(INCENTIVE_TYPE)

    If Len(sType) = 0 Then
        g_clsErrorHandling.RaiseError errGeneralError, "Incentives - Unable to obtain Type"
    End If
    
    GetType = sType
    
    Exit Function
Failed:
    g_clsErrorHandling.RaiseError Err.Number, Err.DESCRIPTION
End Function
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' Function      : GetTypeIndex
' Description   : Returns the Type of Incentive currently selected in numeric form.
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub GetTypeIndex(nIndex As IncentiveType)
    On Error GoTo Failed
    Dim sType As String
    
    sType = GetType
    
    Select Case sType
        Case "Inclusive"
            nIndex = Inclusive
        Case "Exclusive"
            nIndex = Exclusive
        Case Else
            g_clsErrorHandling.RaiseError errGeneralError, "Incentives - Unable to obtain Type from Index"
    End Select
    
    Exit Sub
Failed:
    g_clsErrorHandling.RaiseError Err.Number, Err.DESCRIPTION
End Sub
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' Function      : GetKeyMatchValues
' Description   : Gets the keys associated with the current row of the listview.
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub GetKeyMatchValues(colValues As Collection)
    On Error GoTo Failed
    Dim nListIndex As Integer
    Dim clsPopulateDetails As PopulateDetails
    
    nListIndex = frmProductDetails.lvIncentives.SelectedItem.Index
    
    Set clsPopulateDetails = frmProductDetails.lvIncentives.GetExtra(nListIndex)
    
    If Not clsPopulateDetails Is Nothing Then
        Set colValues = clsPopulateDetails.GetKeyMatchValues()
        If Not colValues Is Nothing Then
            If colValues.Count = 0 Then
                g_clsErrorHandling.RaiseError errGeneralError, "Unable to obtain Incentive Keys"
            End If
        Else
            g_clsErrorHandling.RaiseError errGeneralError, "Unable to obtain Incentive Keys"
        End If
    End If
    
    Exit Sub
Failed:
    g_clsErrorHandling.RaiseError Err.Number, Err.DESCRIPTION
End Sub
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' Function      : ValidateScreenData
' Description   : Validates all fields on the Mortgage Product Incentives tab - called externaly,
'                 and doesn't actually validate any fields on this tab.
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Public Function ValidateScreenData() As Boolean
    Dim bRet As Boolean
    
    bRet = True
    
    ValidateScreenData = bRet
End Function
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' Function      : SetProductKeys
' Description   : Sets the product keys to be used by this tab when retrieving data, if in
'                 edit mode.
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Public Sub SetProductKeys(colValues As Collection)
    On Error GoTo Failed
    
    If colValues Is Nothing Then
        g_clsErrorHandling.RaiseError errKeysEmpty, "AdditionalMortProdParams.SetMortgageParameterKeys"
    End If

    m_sProductCode = colValues(PRODUCT_CODE_KEY)
    m_sProductStartDate = colValues(PRODUCT_START_DATE_KEY)

    Set m_colProductKeys = colValues
    
    Exit Sub
Failed:
    g_clsErrorHandling.RaiseError Err.Number, Err.DESCRIPTION
End Sub
Private Sub Class_Terminate()
    Set m_colIncentiveGUIDs = Nothing
End Sub
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' Function      : SetButtonState
' Description   : Sets the state of the Add/Edit/Delete buttons on the Incentives tab depending
'                 on the state of the Incentives listview.
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub SetButtonState()
    On Error GoTo Failed
    Dim nRecordCount As Long
    Dim bEnableEdit As Boolean
    Dim bEnableDelete As Boolean
    Dim lstItem As ListItem
    
    nRecordCount = frmProductDetails.lvIncentives.ListItems.Count
            
    bEnableEdit = False
    bEnableDelete = False
    
    Set lstItem = frmProductDetails.lvIncentives.SelectedItem
    
    If Not lstItem Is Nothing Then
        bEnableDelete = True
        bEnableEdit = True
    End If
    
    frmProductDetails.cmdIncentivesEdit.Enabled = bEnableEdit
    frmProductDetails.cmdIncentivesDelete.Enabled = bEnableDelete
        
    Exit Sub
Failed:
    g_clsErrorHandling.RaiseError Err.Number, Err.DESCRIPTION
End Sub
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' Function      : InitialiseFields
' Description   : Sets the state of fields and buttons on the Incentives Fees tab.
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub InitialiseFields()
On Error GoTo Failed

    frmProductDetails.cmdIncentivesDelete.Enabled = False
    frmProductDetails.cmdIncentivesEdit.Enabled = False

Exit Sub
Failed:
    g_clsErrorHandling.RaiseError Err.Number, "InitialiseFields:" + Err.DESCRIPTION
End Sub
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' Function      : InitTab
' Description   : Called whenever this tab is switched to, so perform any initialisation required
'                 by the tab when it gets focus.
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Public Sub InitTab()
    On Error GoTo Failed
    
    Exit Sub
Failed:
    g_clsErrorHandling.RaiseError Err.Number, Err.DESCRIPTION
End Sub


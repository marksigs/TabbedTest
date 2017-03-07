VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "MortProdIntRates"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' Class         : MortProdIntRates.cls
' Description   : Contains all processing for the Mortgage Product Interest Rates tab on the Mortgage
'                 Product details screen
' Change history
' Prog      Date        Description
' DJP       03/12/01    SYS2912 SQL Server locking problem.
' DJP       10/12/01    SYS2831 Client variants.
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Option Explicit

' Private data
Private m_colProductKeys As Collection
Private m_clsInterestRates As MortProdIntRatesTable
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' Function      : Class_Initialize
' Description   : Called when this class is created. Create all module data as required.
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub Class_Initialize()
    Set m_clsInterestRates = New MortProdIntRatesTable
End Sub
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' Function      : Initialise
' Description   : Called from the Mortgage Product form to do initialisation of the Mortgage
'                 Product Interest Rates tab. Needs to read the Interest Rates from the database
'                 and perform any validation required.
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Public Sub Initialise(bIsEdit As Boolean)
    On Error GoTo Failed

    InitialiseFields
    SetupListViewHeaders
    PopulateInterestRates
    PopulateListView
    
    frmProductDetails.lvInterestRates.LoadColumnDetails TypeName(Me)

    Exit Sub
Failed:
    g_clsErrorHandling.RaiseError Err.Number, Err.DESCRIPTION
End Sub
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' Function      : PopulateInterestRates
' Description   : Retrieves all Interest Rates for hte current product (but only if product keys
'                 have been supplied).
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub PopulateInterestRates()
    On Error GoTo Failed
    Dim clsIncentives As MortProdIntRatesTable
    Set clsIncentives = New MortProdIntRatesTable
    
    If Not m_colProductKeys Is Nothing Then
        If m_colProductKeys.Count > 0 Then
            m_clsInterestRates.GetInterestRates m_colProductKeys
        End If
    End If
    
    Exit Sub
Failed:
    g_clsErrorHandling.RaiseError Err.Number, Err.DESCRIPTION
End Sub
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' Function      : PopulateListView
' Description   : RePopulates the Interest Rate listview
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub PopulateListView()
    On Error GoTo Failed
    
    g_clsFormProcessing.PopulateFromRecordset frmProductDetails.lvInterestRates, m_clsInterestRates
    
    ' Clear any selections we may have had.
    Set frmProductDetails.lvInterestRates.SelectedItem = Nothing
    
    Exit Sub
Failed:
    g_clsErrorHandling.RaiseError Err.Number, Err.DESCRIPTION
End Sub
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' Function      : SetupListViewHeaders
' Description   : Sets the listview column headings titles and sizes.
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub SetupListViewHeaders()
    On Error GoTo Failed
    Dim headers As New Collection
    Dim lvHeaders As listViewAccess

    lvHeaders.nWidth = 15
    lvHeaders.sName = "Sequence Number"
    headers.Add lvHeaders

    lvHeaders.nWidth = 25
    lvHeaders.sName = "Type"
    headers.Add lvHeaders

    lvHeaders.nWidth = 15
    lvHeaders.sName = "Period"
    headers.Add lvHeaders

    lvHeaders.nWidth = 20
    lvHeaders.sName = "End Date"
    headers.Add lvHeaders

    lvHeaders.nWidth = 20
    lvHeaders.sName = "Rate"
    headers.Add lvHeaders

    lvHeaders.nWidth = 20
    lvHeaders.sName = "Ceiling"
    headers.Add lvHeaders

    lvHeaders.nWidth = 20
    lvHeaders.sName = "Floored"
    headers.Add lvHeaders

    frmProductDetails.lvInterestRates.AddHeadings headers
    
    Exit Sub
Failed:
    g_clsErrorHandling.RaiseError Err.Number, Err.DESCRIPTION
End Sub
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' Function      : Add
' Description   : Called when the user presses the Add button on the Product Details form,
'                 Interest Rates tab. Show the Interest Rates form.
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Public Sub Add()
    On Error GoTo Failed

    frmEditInterestRates.SetIsEdit False
    frmEditInterestRates.SetTableClass m_clsInterestRates
    frmEditInterestRates.Show vbModal, frmProductDetails

    PopulateListView
    SetButtonState
    Unload frmEditInterestRates
    
    Exit Sub
Failed:
    g_clsErrorHandling.DisplayError
End Sub
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' Function      : Edit
' Description   : Called when the user presses the Edit button on the Product Details form,
'                 Interest Rates tab. Edit the Interest Rate - load the Interest Rates form.
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Public Sub Edit()
    On Error GoTo Failed
    Dim bRet As Boolean
    Dim colMatchValues As New Collection
    Dim clsTableAccess As TableAccess

    ' Get the keys for the currently selected Interest Rate
    bRet = GetKeyMatchValues(colMatchValues)
    
    If bRet = True Then
        Set clsTableAccess = m_clsInterestRates
        clsTableAccess.SetKeyMatchValues colMatchValues
        
        frmEditInterestRates.SetTableClass clsTableAccess
        frmEditInterestRates.SetIsEdit
        frmEditInterestRates.Show vbModal, frmProductDetails
        
        If frmEditInterestRates.GetReturnCode = MSGSuccess Then
            PopulateListView
            SetButtonState
        End If
        
        Unload frmEditInterestRates
    End If
    
    Exit Sub
Failed:
    g_clsErrorHandling.RaiseError Err.Number, Err.DESCRIPTION
End Sub
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' Function      : Delete
' Description   : Called when the user presses the Delete button on the Product Details form,
'                 Interest Rates tab. Delete the Interest Rate - load the Interest Rates form.
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Public Sub Delete()
    Dim bRet As Boolean
    Dim nResponse As Integer
    Dim colMatchValues As New Collection

    On Error GoTo Failed

    If Not frmProductDetails.lvInterestRates.SelectedItem Is Nothing Then
        nResponse = MsgBox("Delete the selected record?", vbQuestion + vbYesNo)
    
        If nResponse = vbYes Then
            bRet = GetKeyMatchValues(colMatchValues)
    
            If bRet = True Then
                TableAccess(m_clsInterestRates).SetKeyMatchValues colMatchValues
                TableAccess(m_clsInterestRates).ApplyFilter
                TableAccess(m_clsInterestRates).DeleteAllRows
                
                TableAccess(m_clsInterestRates).CancelFilter
                PopulateListView
                SetButtonState
            End If
        End If
    End If
    
    Exit Sub
Failed:
    g_clsErrorHandling.RaiseError Err.Number, Err.DESCRIPTION
End Sub
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' Function      : SaveScreenData
' Description   : Saves all fields on the Interest Rates tab
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Public Sub SaveScreenData()
    On Error GoTo Failed
    
    ' Write the product code and start date to the Interest rate records (if adding, there won't be a Product
    ' Code and Start Date)
    SetRateProductKeys
    
    ' Update the record.
    If TableAccess(m_clsInterestRates).RecordCount > 0 Then
        TableAccess(m_clsInterestRates).Update
    End If
    
    Exit Sub
Failed:
    g_clsErrorHandling.RaiseError Err.Number, Err.DESCRIPTION
End Sub
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' Function      : SetRateProductKeys
' Description   : Writes the Produt Code and Start Date to all Interest Rate records.
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub SetRateProductKeys()
    On Error GoTo Failed
    Dim clsBanded As BandedTable
    
    If TableAccess(m_clsInterestRates).RecordCount > 0 Then
        Set clsBanded = m_clsInterestRates
        clsBanded.SetUpdateValues m_colProductKeys
        clsBanded.SetUpdateSets
        clsBanded.DoUpdateSets
    End If
    
    Exit Sub
Failed:
    g_clsErrorHandling.RaiseError Err.Number, Err.DESCRIPTION
End Sub
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' Function      : GetKeyMatchValues
' Description   : Gets the keys associated with the current row of the listview.
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Function GetKeyMatchValues(colValues As Collection) As Boolean
    On Error GoTo Failed
    Dim bRet As Boolean
    Dim nListIndex As Integer
    Dim clsPopulateDetails  As PopulateDetails
    
    bRet = False
    nListIndex = frmProductDetails.lvInterestRates.SelectedItem.Index

    Set clsPopulateDetails = frmProductDetails.lvInterestRates.GetExtra(nListIndex)
    
    If Not clsPopulateDetails Is Nothing Then
        Set colValues = clsPopulateDetails.GetKeyMatchValues()
        If Not colValues Is Nothing Then
            If colValues.Count > 0 Then
                bRet = True
            End If
        End If
    End If
    
    GetKeyMatchValues = bRet
    
    Exit Function
Failed:
    g_clsErrorHandling.RaiseError Err.Number, Err.DESCRIPTION
End Function
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' Function      : SetScreenFields
' Description   : Populates all fields on the Mortgage Product Interest Rates tab - called externaly,
'                 and doesn't actually populate any fields on this tab.
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Public Sub SetScreenFields()
    On Error GoTo Failed

    Exit Sub
Failed:
    g_clsErrorHandling.RaiseError Err.Number, Err.DESCRIPTION
End Sub
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' Function      : ValidateScreenData
' Description   : Validates all fields on the Mortgage Product Interest Rates tab - called externaly,
'                 and doesn't actually validate any fields on this tab.
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Public Function ValidateScreenData() As Boolean
    On Error GoTo Failed
    Dim bRet As Boolean
    
    bRet = True
    
    ValidateScreenData = bRet
    
    Exit Function
Failed:
    g_clsErrorHandling.RaiseError Err.Number, Err.DESCRIPTION
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
    Set m_colProductKeys = colValues
    Exit Sub
Failed:
    g_clsErrorHandling.RaiseError Err.Number, Err.DESCRIPTION
End Sub
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' Function      : SetButtonState
' Description   : Sets the state of the Add/Edit/Delete buttons on the Interest Rates tab depending
'                 on the state of the Interest Rates listview.
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub SetButtonState()
    On Error GoTo Failed
    Dim nRecordCount As Long
    Dim bEnableEdit As Boolean
    Dim bEnableDelete As Boolean
    Dim lstItem As ListItem
    
    nRecordCount = TableAccess(m_clsInterestRates).RecordCount
            
    bEnableEdit = False
    bEnableDelete = False
    
    Set lstItem = frmProductDetails.lvInterestRates.SelectedItem
    
    If Not lstItem Is Nothing Then
        bEnableDelete = True
        bEnableEdit = True
    End If
    
    frmProductDetails.cmdInterestRateEdit.Enabled = bEnableEdit
    frmProductDetails.cmdInterestRateDelete.Enabled = bEnableDelete
        
    Exit Sub
Failed:
    g_clsErrorHandling.RaiseError Err.Number, Err.DESCRIPTION
End Sub
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' Function      : InitialiseFields
' Description   : Sets the state of fields and buttons on the Interest Rates tab.
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub InitialiseFields()
On Error GoTo Failed

    frmProductDetails.cmdInterestRateEdit.Enabled = False
    frmProductDetails.cmdInterestRateDelete.Enabled = False

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


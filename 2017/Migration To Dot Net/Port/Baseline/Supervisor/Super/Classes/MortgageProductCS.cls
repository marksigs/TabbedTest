VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "MortgageProductCS"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' Class         : MortgageProductCS
' Description   : Client specific version of the MortgageProduct class. This class will be called by
'                 a form or controlling class and any client specific code can be put in here for
'                 each client. Each method can a) override the base code, b) delegate to the base
'                 code or c) overried and delegate.
' Change history
' Prog      Date        Description
' DJP       22/10/01    Created.
' DJP       20/11/01    SYS2831/SYS2912 Support client variants & SQL Server locking problem.
' DJP       14/02/02    SYS4094 Set Mortgage Product form to be used for Base Rates.

' BMIDS Change history
' Prog      Date        Description
' GD        22/05/02    BMIDS0012 : Methods changed : MortgageProduct_CreateTabs(),MortgageProduct_InitialiseTab()
' DJP       06/03/03    BM0423 Validate the mortgage product when OK is pressed.
' MC        24/05/2004  BMIDS775/REG016/018 AdditionalFeaturedPressed Event implemented
'                       MortgageProduct Interface has been extended.
' HMA       20/12/2004  BMIDS959  Remove ValidateRedemption

' EPSOM Change history
' Prog      Date        Description
' TW        09/10/2006  EP2_7 Added processing for Additional Borrowing Fee and Credit Limit Increase Fee
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''

Option Explicit

' Base class
Implements MortgageProduct

' Private data
Private m_clsMortgageProduct As MortgageProduct
Private m_colTabs As Collection
Private m_colKeys As Collection

'*=[MC]BMIDS775/REG016/018
Private Sub MortgageProduct_AdditionalFeaturesPressed(Optional ByVal vstrMortProdCode As String, Optional ByVal vstrStartDt As String)
    m_clsMortgageProduct.AdditionalFeaturesPressed vstrMortProdCode, vstrStartDt
End Sub

'*=[MC]BMIDS775/REG016/018
'Private Sub MortgageProduct_AdditionalFeaturesPressed()
'    m_clsMortgageProduct.AdditionalFeaturesPressed
'End Sub


Private Function MortgageProduct_GetKeyMatchValues() As Collection
    Set MortgageProduct_GetKeyMatchValues = m_clsMortgageProduct.GetKeyMatchValues
End Function

Private Function MortgageProduct_GetTabsCollection() As Collection
    Set MortgageProduct_GetTabsCollection = m_clsMortgageProduct.GetTabsCollection
End Function

Private Sub MortgageProduct_Initialise()
    On Error GoTo Failed
    
    m_clsMortgageProduct.Initialise
    
    Exit Sub
Failed:
    g_clsErrorHandling.RaiseError Err.Number, Err.DESCRIPTION
End Sub

Private Sub MortgageProduct_Show()
    On Error GoTo Failed
    
    ' Perform MortgageProduct load functionality first
    MortgageProduct_Load
    
    m_clsMortgageProduct.Show
    
    Exit Sub
Failed:
    g_clsErrorHandling.RaiseError Err.Number, Err.DESCRIPTION
End Sub

Private Sub MortgageProduct_Load()
    On Error GoTo Failed
    
    ' Delete to Core method
    m_clsMortgageProduct.Load
    
    ' Create tab control classes
    MortgageProduct_CreateTabs
    
    Exit Sub
Failed:
    g_clsErrorHandling.RaiseError Err.Number, Err.DESCRIPTION
End Sub

Private Sub MortgageProduct_CreateTabs()

'GD BMIDS0012 - Addition of tab
    On Error GoTo Failed
    Dim bEdit As Boolean
    Dim clsBaseRates As MortProdBaseRates
    
    bEdit = m_clsMortgageProduct.IsEdit
    
    ' Delete to Core method
    m_clsMortgageProduct.CreateTabs
    
    ' Set the tabs collection for this class to use, now that Core has created all required tabs
    Set m_colTabs = m_clsMortgageProduct.GetTabsCollection
    
    ' ... Create any client specific tabs here ...
    'GD BMIDS START
    Dim clsMortProdProdCond As MortProdProdCond
    Set clsMortProdProdCond = New MortProdProdCond
    
    m_colTabs.Add clsMortProdProdCond
    'END GD BMIDS
    ' Set the Base Rates form
    Set clsBaseRates = m_colTabs(mortAreaTabBaseRates)
    clsBaseRates.SetForm frmProductDetails
    
    If bEdit = True Then
        InitialiseEdit
        ' Initialise any client specific editing
    Else
        ' Need to be able to pass the MortgageProduct code across to the LegalFee screen so it can be used
        ' to a) display the MortgageProduct code on the screen in an edit, and also when the user presses
        ' ok we need to use the MortgageProduct code the user had entered as a key for the LegalFeeSet
        ' table.
        
        InitialiseAdd
        ' Initialise any client specific Adding
    
    End If
    
    ' Now initialise all tabs (this will include any client tabs added)
    MortgageProduct_InitialiseTabs
    MortgageProduct_PopulateScreenFields
    
    Exit Sub
Failed:
    g_clsErrorHandling.RaiseError Err.Number, Err.DESCRIPTION
End Sub


Private Sub InitialiseEdit()
    On Error GoTo Failed
        
    ' Delegate to the Core method
    m_clsMortgageProduct.Initialise
    
    ' Get the Core MortgageProduct keys for use with client versions
    Set m_colKeys = MortgageProduct_GetKeyMatchValues

    ' ... Do any client specific processing here ...
    
    Exit Sub
Failed:
    g_clsErrorHandling.RaiseError Err.Number, Err.DESCRIPTION
End Sub

Private Sub InitialiseAdd()
    On Error GoTo Failed
    
    ' Delegate to the Core method
    'm_clsMortgageProduct.Initialise

    ' ... Do any client specific processing here ...
    
    Exit Sub
Failed:
    g_clsErrorHandling.RaiseError Err.Number, Err.DESCRIPTION
End Sub

Private Sub MortgageProduct_InitialiseTabs()
    On Error GoTo Failed
    
    ' Delegate to the Core method
    m_clsMortgageProduct.InitialiseTabs

    
    ' ... Put any client specific code here ...

    
    Exit Sub
Failed:
    g_clsErrorHandling.RaiseError Err.Number, Err.DESCRIPTION
End Sub

Private Sub MortgageProduct_Add(nItem As ProductArea)
    m_clsMortgageProduct.Add nItem
End Sub

Private Sub MortgageProduct_Cancel()
    m_clsMortgageProduct.Cancel
End Sub

Private Sub MortgageProduct_Delete(nItem As ProductArea)
    m_clsMortgageProduct.Delete nItem
End Sub

Private Sub MortgageProduct_DeselectItem(nItem As ProductArea)
    m_clsMortgageProduct.DeselectItem nItem
End Sub

Private Sub MortgageProduct_Edit(nItem As ProductArea)
    m_clsMortgageProduct.Edit nItem
End Sub

Private Function MortgageProduct_GetReturnCode() As MSGReturnCode
    MortgageProduct_GetReturnCode = m_clsMortgageProduct.GetReturnCode
End Function


Private Sub MortgageProduct_InitialiseTab()
'GD BMIDS0012
    On Error GoTo Failed
    Dim nTab As Integer
    m_clsMortgageProduct.InitialiseTab
    If Not m_colTabs Is Nothing Then
        nTab = frmProductDetails.SSTab1.Tab + 1
        Select Case nTab
            Case mortProdAppEligibility
                m_colTabs(mortAreaTypeOfApp).InitTab
                m_colTabs(mortAreaPurpOfLoan).InitTab
                m_colTabs(mortAreaTypeOfBuyer).InitTab
        End Select
    End If
    Exit Sub
Failed:
    g_clsErrorHandling.RaiseError Err.Number, Err.DESCRIPTION
End Sub


Private Function MortgageProduct_IsEdit() As Boolean
    MortgageProduct_IsEdit = m_clsMortgageProduct.IsEdit
End Function

Private Sub MortgageProduct_LenderChange()
    m_clsMortgageProduct.LenderChange
End Sub

Private Sub MortgageProduct_OK()
    On Error GoTo Failed
    Dim bSuccess As Boolean
    ' First, delegate to Core for Validation
    
    ' Delegate to the Core Validation method first.
    bSuccess = m_clsMortgageProduct.ValidateScreenData
    
    If bSuccess Then
        ' Delegate to Core to save existing data
        m_clsMortgageProduct.SaveScreenData
        
        ' ...Do any client specific data saving here...
        m_clsMortgageProduct.OK
    End If
    
    Exit Sub
Failed:
    g_clsErrorHandling.RaiseError Err.Number, Err.DESCRIPTION
End Sub

Private Sub MortgageProduct_PopulateScreenFields()
    On Error GoTo Failed
    
    m_clsMortgageProduct.PopulateScreenFields
    
    Exit Sub
Failed:
    g_clsErrorHandling.RaiseError Err.Number, Err.DESCRIPTION
End Sub

Private Sub MortgageProduct_ProductGroupingPressed()
    m_clsMortgageProduct.ProductGroupingPressed
End Sub

Private Sub MortgageProduct_SaveScreenData()
    On Error GoTo Failed
    
    ' First, delegate to the Core method
    m_clsMortgageProduct.SaveScreenData

    ' Now do any client saves

    Exit Sub
Failed:
    g_clsErrorHandling.RaiseError Err.Number, Err.DESCRIPTION
    

End Sub

Private Sub MortgageProduct_SaveSets()
    m_clsMortgageProduct.SaveSets
End Sub

Private Sub MortgageProduct_SelectItem(nItem As ProductArea)
    m_clsMortgageProduct.SelectItem nItem
End Sub

Private Sub MortgageProduct_SetActiveTab(Optional nTab As ProductArea)
    m_clsMortgageProduct.SetActiveTab nTab
End Sub

Private Sub MortgageProduct_SetEditState()
    m_clsMortgageProduct.SetEditState
End Sub

Private Sub MortgageProduct_SetIsEdit(Optional bIsEdit As Boolean = True)
    m_clsMortgageProduct.SetIsEdit bIsEdit
End Sub

Private Sub MortgageProduct_SetKeys(colKeys As Collection)
    m_clsMortgageProduct.SetKeys colKeys
End Sub

Private Sub MortgageProduct_SetupListViewFields()
    m_clsMortgageProduct.SetupListViewFields
End Sub

Private Sub MortgageProduct_UnloadForm()
    m_clsMortgageProduct.UnloadForm
End Sub

' BMIDS959  Remove MortgageProduct_ValidateRedemption

Private Function MortgageProduct_ValidateScreenData() As Boolean
    MortgageProduct_ValidateScreenData = m_clsMortgageProduct.ValidateScreenData
End Function

Private Sub Class_Initialize()
    On Error GoTo Failed
    
    Set m_clsMortgageProduct = New MortgageProduct
    frmProductDetails.SetMortgageProduct Me
    Exit Sub
Failed:
    g_clsErrorHandling.RaiseError Err.Number, Err.DESCRIPTION
End Sub
' DJP BM0423
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' Function      : ValidateProduct
' Description   : Calls the ValidateProduct Stored Procedure.
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub ValidateProduct()
    On Error GoTo Failed
    Dim clsMortProd As MortgageProductTable
    Set clsMortProd = New MortgageProductTable
    
    TableAccess(clsMortProd).SetKeyMatchValues m_clsMortgageProduct.GetKeyMatchValues
    
    clsMortProd.ValidateProduct
    
    Exit Sub
Failed:
    g_clsErrorHandling.RaiseError Err.Number, Err.DESCRIPTION
End Sub

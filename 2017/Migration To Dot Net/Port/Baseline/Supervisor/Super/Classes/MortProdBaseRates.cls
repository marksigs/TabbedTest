VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "MortProdBaseRates"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' Class         : MortProdBaseRates
' Description   : Provides functionality for the Base Rates tab on Mortgage Product. Needs to
'                 update the base rate set on the product, so the Mortgage Product table class
'                 is passed in for this to happen.
' Change history
' Prog      Date        Description
' DJP       22/06/01    SYS2381 SQL Server port, and tidy up
' DJP       20/11/01    SYS2912 SQL Server locking problem.
' DJP       10/12/01    SYS2831 Client variants.
' DJP       14/02/02    SYS4094 Allow any form to be used with this class.
' DJP       14/02/02    SYS4095 Don't disable Base Rate DeSelect by default.
' DJP       14/02/02    SYS4095 Reset BaseFeeSet when DeSelected.
' STB       08/03/01    SYS4238 Altered Rate column to 'Rate Diff'.
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Option Explicit

' Constants
Private Const SET_KEY As Integer = 1
Private Const DATE_KEY As Integer = 2

' Private data
Private m_bIsEdit As Boolean
Private m_sBaseRateSet As String
Private m_clsBaseRates As BaseRateTable

' Form to be used by this class
Private m_frmInterestRate As Form
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' Function      : Class_Initialize
' Description   : Called when this class is created. Create all module data as required.
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''

'BMIDS Change history
'Prog      Date         Description
'GD        16/05/2002   BMIDS00009. Amended :Initialise,SetScreenFields.
'                                   Added   :None.


Private Sub Class_Initialize()
    On Error GoTo Failed
    
    Set m_clsBaseRates = New BaseRateTable
    
    Exit Sub
Failed:
    g_clsErrorHandling.RaiseError Err.Number, Err.DESCRIPTION
End Sub


Public Sub Initialise(bIsEdit As Boolean)
'GD        16/05/2002   BMIDS00009
    On Error GoTo Failed
    Dim clsTableAccess As TableAccess
    
    m_bIsEdit = bIsEdit
    'GD BMIDS
    'InitialiseFields
    'SetupListViewHeaders
    'PopulateListView
    
    'm_frmInterestRate.lvBaseRates.LoadColumnDetails TypeName(Me)
    
    Exit Sub
Failed:
    g_clsErrorHandling.RaiseError Err.Number, Err.DESCRIPTION
End Sub


Public Sub SetForm(frmToUse As Form)
    On Error GoTo Failed
    
    If frmToUse Is Nothing Then
        g_clsErrorHandling.RaiseError errGeneralError, TypeName(Me) & ".SetForm Form is empty"
    End If
    
    Set m_frmInterestRate = frmToUse
    
    Exit Sub
Failed:
    g_clsErrorHandling.RaiseError Err.Number, Err.DESCRIPTION
End Sub

Public Function GetForm() As Form
    On Error GoTo Failed
    
    Set GetForm = m_frmInterestRate
    
    Exit Function
Failed:
    g_clsErrorHandling.RaiseError Err.Number, Err.DESCRIPTION
End Function

Private Sub PopulateListView()
    On Error GoTo Failed
    m_clsBaseRates.SetFindNewestBaseRateSet
    
    TableAccess(m_clsBaseRates).GetTableData POPULATE_FIRST_BAND
        
    TableAccess(m_clsBaseRates).ValidateData
    g_clsFormProcessing.PopulateFromRecordset m_frmInterestRate.lvBaseRates, TableAccess(m_clsBaseRates)
    
    Exit Sub
Failed:
    g_clsErrorHandling.RaiseError Err.Number, Err.DESCRIPTION
End Sub

Private Sub SetupListViewHeaders()
    Dim headers As New Collection
    Dim lvHeaders As listViewAccess
    
    lvHeaders.nWidth = 10
    lvHeaders.sName = "Set No."
    headers.Add lvHeaders
    
    lvHeaders.nWidth = 20
    lvHeaders.sName = "Start Date"
    headers.Add lvHeaders
    
    lvHeaders.nWidth = 25
    lvHeaders.sName = "Max. Loan Amount"
    headers.Add lvHeaders
    
    lvHeaders.nWidth = 25
    lvHeaders.sName = "Max. LTV"
    headers.Add lvHeaders
    
    lvHeaders.nWidth = 20
    lvHeaders.sName = "Rate Diff"
    headers.Add lvHeaders
    
    m_frmInterestRate.lvBaseRates.AddHeadings headers
End Sub

Public Sub SetScreenFields()
'GD        16/05/2002   BMIDS00009
    On Error GoTo Failed
    'GD BMIDS
    'm_frmInterestRate.txtBaseRateSelected.Text = m_sBaseRateSet

    Exit Sub
Failed:
    g_clsErrorHandling.RaiseError Err.Number, Err.DESCRIPTION
End Sub


Public Sub SaveScreenData()
    On Error GoTo Failed
    Exit Sub
Failed:
    g_clsErrorHandling.RaiseError Err.Number, Err.DESCRIPTION
End Sub

Private Sub GetKeyMatchValues(colValues As Collection)
    On Error GoTo Failed
    Dim nListIndex As Integer
    Dim clsPopulateDetails As PopulateDetails
    
    nListIndex = m_frmInterestRate.lvBaseRates.SelectedItem.Index
    
    Set clsPopulateDetails = m_frmInterestRate.lvBaseRates.GetExtra(nListIndex)
    
    If Not clsPopulateDetails Is Nothing Then
        Set colValues = clsPopulateDetails.GetKeyMatchValues()
        If Not colValues Is Nothing Then
            If colValues.Count = 0 Then
                g_clsErrorHandling.RaiseError errGeneralError, "Unable to obtain Base Rate Keys"
            End If
        Else
            g_clsErrorHandling.RaiseError errGeneralError, "Unable to obtain Base Rate Keys"
        End If
    End If
    
    Exit Sub
Failed:
    g_clsErrorHandling.RaiseError Err.Number, Err.DESCRIPTION
End Sub

Public Function ValidateScreenData() As Boolean
    On Error GoTo Failed
    
    ValidateScreenData = True
    
    Exit Function
Failed:
    g_clsErrorHandling.RaiseError Err.Number, Err.DESCRIPTION
End Function

Public Sub DeSelect()
    On Error GoTo Failed
    
    m_frmInterestRate.txtBaseRateSelected.Text = ""
    m_frmInterestRate.cmdBaseRateDeSelect.Enabled = False
    
    ' Reset BaseRateSet
    m_sBaseRateSet = ""
    
    Exit Sub
Failed:
    g_clsErrorHandling.RaiseError Err.Number, Err.DESCRIPTION
End Sub

Public Sub Selected()
    On Error GoTo Failed
    Dim colValues As Collection
    Dim bRet As Boolean
    Set colValues = New Collection
    
    GetKeyMatchValues colValues
    
    If colValues.Count >= SET_KEY Then
        m_sBaseRateSet = colValues(SET_KEY)
        m_frmInterestRate.txtBaseRateSelected.Text = m_sBaseRateSet
        m_frmInterestRate.cmdBaseRateDeSelect.Enabled = True
    End If
    
    Exit Sub
Failed:
    g_clsErrorHandling.RaiseError Err.Number, Err.DESCRIPTION
End Sub

Public Sub SetBaseRateSet(sRateSet As String)
    On Error GoTo Failed
    
    m_sBaseRateSet = sRateSet
    
    Exit Sub
Failed:
    g_clsErrorHandling.RaiseError Err.Number, Err.DESCRIPTION
End Sub

Public Function GetBaseRateSet() As String
    On Error GoTo Failed
    
    GetBaseRateSet = m_sBaseRateSet
    
    Exit Function
Failed:
    g_clsErrorHandling.RaiseError Err.Number, Err.DESCRIPTION
End Function

Public Sub SetProductKeys(colValues As Collection)
    On Error GoTo Failed
    
    Exit Sub
Failed:
    g_clsErrorHandling.RaiseError Err.Number, Err.DESCRIPTION
End Sub

Private Sub InitialiseFields()
On Error GoTo Failed

    m_frmInterestRate.cmdBaseRateSelect.Enabled = False

Exit Sub
Failed:
    g_clsErrorHandling.RaiseError Err.Number, "InitialiseFields:" + Err.DESCRIPTION
End Sub

Public Sub InitTab()
    On Error GoTo Failed
    
    Exit Sub
Failed:
    g_clsErrorHandling.RaiseError Err.Number, Err.DESCRIPTION
End Sub


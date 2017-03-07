VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "MortProdBaseRatesCS"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' Class         : MortProdBaseRatesCS
' Description   : Client specific version of the MortProdBaseRates class. This class will be called by
'                 a form or controlling class and any client specific code can be put in here for
'                 each client. Each method can a) override the base code, b) delegate to the base
'                 code or c) overried and delegate.
' Change history
' Prog      Date        Description
' DJP       10/12/01    SYS2831 Support client variants
' DJP       14/02/02    SYS4094 Allow any form to be used with this class.
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Option Explicit

' Base class
Implements MortProdBaseRates

' Private data
Private m_clsMortProdBaseRates As MortProdBaseRates
Private Sub Class_Initialize()
    On Error GoTo Failed
    
    Set m_clsMortProdBaseRates = New MortProdBaseRates
    
    Exit Sub
Failed:
    g_clsErrorHandling.RaiseError Err.Number, Err.DESCRIPTION
End Sub
Private Sub MortProdBaseRates_DeSelect()
    m_clsMortProdBaseRates.DeSelect
End Sub
Private Function MortProdBaseRates_GetBaseRateSet() As String
    MortProdBaseRates_GetBaseRateSet = m_clsMortProdBaseRates.GetBaseRateSet
End Function
Private Function MortProdBaseRates_GetForm() As Form
    Set MortProdBaseRates_GetForm = m_clsMortProdBaseRates.GetForm
End Function
Private Sub MortProdBaseRates_Initialise(bIsEdit As Boolean)
    m_clsMortProdBaseRates.Initialise bIsEdit
End Sub
Private Sub MortProdBaseRates_InitTab()
    m_clsMortProdBaseRates.InitTab
End Sub
Private Sub MortProdBaseRates_SaveScreenData()
    m_clsMortProdBaseRates.SaveScreenData
End Sub
Private Sub MortProdBaseRates_Selected()
    m_clsMortProdBaseRates.Selected
End Sub
Private Sub MortProdBaseRates_SetBaseRateSet(sRateSet As String)
    m_clsMortProdBaseRates.SetBaseRateSet sRateSet
End Sub
Private Sub MortProdBaseRates_SetForm(frmToUse As Form)
    m_clsMortProdBaseRates.SetForm frmToUse
End Sub
Private Sub MortProdBaseRates_SetProductKeys(colValues As Collection)
    m_clsMortProdBaseRates.SetProductKeys colValues
End Sub
Private Sub MortProdBaseRates_SetScreenFields()
    m_clsMortProdBaseRates.SetScreenFields
End Sub
Private Function MortProdBaseRates_ValidateScreenData() As Boolean
    MortProdBaseRates_ValidateScreenData = m_clsMortProdBaseRates.ValidateScreenData
End Function

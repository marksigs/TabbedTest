VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "MortProdIntRatesCS"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' Class         : MortProdIntRatesCS
' Description   : Client specific version of the MortProdIntRates class. This class will be called by
'                 a form or controlling class and any client specific code can be put in here for
'                 each client. Each method can a) override the base code, b) delegate to the base
'                 code or c) overried and delegate.
' Change history
' Prog      Date        Description
' DJP       10/12/01    SYS2831 Support client variants
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Option Explicit

' Base class
Implements MortProdIntRates

' Private data
Private m_clsMortProdIntRates As MortProdIntRates

Private Sub Class_Initialize()
    On Error GoTo Failed
    
    Set m_clsMortProdIntRates = New MortProdIntRates
    
    Exit Sub
Failed:
    g_clsErrorHandling.RaiseError Err.Number, Err.DESCRIPTION
End Sub

Private Sub MortProdIntRates_Add()
    m_clsMortProdIntRates.Add
End Sub

Private Sub MortProdIntRates_Delete()
    m_clsMortProdIntRates.Delete
End Sub

Private Sub MortProdIntRates_Edit()
    m_clsMortProdIntRates.Edit
End Sub

Private Sub MortProdIntRates_Initialise(bIsEdit As Boolean)
    m_clsMortProdIntRates.Initialise bIsEdit
End Sub

Private Sub MortProdIntRates_InitTab()
    m_clsMortProdIntRates.InitTab
End Sub

Private Sub MortProdIntRates_SaveScreenData()
    m_clsMortProdIntRates.SaveScreenData
End Sub

Private Sub MortProdIntRates_SetProductKeys(colValues As Collection)
    m_clsMortProdIntRates.SetProductKeys colValues
End Sub

Private Sub MortProdIntRates_SetScreenFields()
    m_clsMortProdIntRates.SetScreenFields
End Sub

Private Function MortProdIntRates_ValidateScreenData() As Boolean
    MortProdIntRates_ValidateScreenData = m_clsMortProdIntRates.ValidateScreenData
End Function


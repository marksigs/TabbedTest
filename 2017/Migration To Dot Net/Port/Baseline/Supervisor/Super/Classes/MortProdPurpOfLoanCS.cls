VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "MortProdPurpOfLoanCS"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' Class         : MortProdPurpOfLoanCS
' Description   : Client specific version of the MortProdPurpOfLoan class. This class will be called by
'                 a form or controlling class and any client specific code can be put in here for
'                 each client. Each method can a) override the base code, b) delegate to the base
'                 code or c) overried and delegate.
' Change history
' Prog      Date        Description
' DJP       10/12/01    SYS2831 Support client variants
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Option Explicit
' Base class
Implements MortProdPurpOfLoan

' Privata data
Private m_clsMortProdPurpOfLoan As MortProdPurpOfLoan
Private Sub Class_Initialize()
    Set m_clsMortProdPurpOfLoan = New MortProdPurpOfLoan
End Sub
Private Sub MortProdPurpOfLoan_DoUpdates()
    m_clsMortProdPurpOfLoan.DoUpdates
End Sub
Private Sub MortProdPurpOfLoan_Initialise(bIsEdit As Boolean)
    m_clsMortProdPurpOfLoan.Initialise bIsEdit
End Sub
Private Sub MortProdPurpOfLoan_InitTab()
    m_clsMortProdPurpOfLoan.InitTab
End Sub
Private Sub MortProdPurpOfLoan_SaveScreenData()
    m_clsMortProdPurpOfLoan.SaveScreenData
End Sub
Private Sub MortProdPurpOfLoan_SetProductKeys(colValues As Collection)
    m_clsMortProdPurpOfLoan.SetProductKeys colValues
End Sub
Private Sub MortProdPurpOfLoan_SetScreenFields()
    m_clsMortProdPurpOfLoan.SetScreenFields
End Sub
Private Function MortProdPurpOfLoan_ValidateScreenData() As Variant
    MortProdPurpOfLoan_ValidateScreenData = m_clsMortProdPurpOfLoan.ValidateScreenData
End Function

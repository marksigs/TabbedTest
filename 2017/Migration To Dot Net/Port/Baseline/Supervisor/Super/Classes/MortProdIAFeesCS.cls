VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "MortProdIAFeesCS"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'*******************************************************************************************************
' Class         : MortProdIAFeesCS
' Description   : Client specific version of the MortProdIAFees class. This class will be called by
'                 a form or controlling class and any client specific code can be put in here for
'                 each client. Each method can a) override the base code, b) delegate to the base
'                 code or c) overried and delegate.
' Change history
' Prog      Date        Description
'   MC      02/06/2004  BMIDS763    BMIDS763 - CC075 - FEESETS Insurance Admin Fee Set CS
'*******************************************************************************************************
Option Explicit

' Base class
Implements MortProdIAFees

' Private data
Private m_clsMortProdIAFees As MortProdIAFees
Private Sub Class_Initialize()
    Set m_clsMortProdIAFees = New MortProdIAFees
End Sub
Private Sub MortProdIAFees_DeSelect()
    m_clsMortProdIAFees.DeSelect
End Sub
Private Function MortProdIAFees_GetIAFeeSet() As Variant
    MortProdIAFees_GetIAFeeSet = m_clsMortProdIAFees.GetIAFeeSet
End Function
Private Sub MortProdIAFees_Initialise(bIsEdit As Boolean)
    m_clsMortProdIAFees.Initialise bIsEdit
End Sub
Private Sub MortProdIAFees_InitTab()
    m_clsMortProdIAFees.InitTab
End Sub
Private Sub MortProdIAFees_SaveScreenData()
    m_clsMortProdIAFees.SaveScreenData
End Sub
Private Sub MortProdIAFees_Selected()
    m_clsMortProdIAFees.Selected
End Sub
Private Sub MortProdIAFees_SetIAFeeSet(sFeeSet As String)
    m_clsMortProdIAFees.SetIAFeeSet sFeeSet
End Sub
Private Sub MortProdIAFees_SetProductKeys(colValues As Collection)
    m_clsMortProdIAFees.SetProductKeys colValues
End Sub
Private Sub MortProdIAFees_SetScreenFields()
    m_clsMortProdIAFees.SetScreenFields
End Sub
Private Function MortProdIAFees_ValidateScreenData() As Boolean
    MortProdIAFees_ValidateScreenData = m_clsMortProdIAFees.ValidateScreenData
End Function




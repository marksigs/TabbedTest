VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "MortProdTTFeesCS"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'*******************************************************************************************************
' Class         : MortProdTTFeesCS
' Description   : Client specific version of the MortProdTTFees class. This class will be called by
'                 a form or controlling class and any client specific code can be put in here for
'                 each client. Each method can a) override the base code, b) delegate to the base
'                 code or c) overried and delegate.
' Change history
' Prog      Date        Description
'   MC      02/06/2004  BMIDS763    BMIDS763 - CC075 - FEESETS New Class refs added
'*******************************************************************************************************
Option Explicit

' Base class
Implements MortProdTTFees

' Private data
Private m_clsMortProdTTFees As MortProdTTFees
Private Sub Class_Initialize()
    Set m_clsMortProdTTFees = New MortProdTTFees
End Sub
Private Sub MortProdTTFees_DeSelect()
    m_clsMortProdTTFees.DeSelect
End Sub
Private Function MortProdTTFees_GetTTFeeSet() As Variant
    MortProdTTFees_GetTTFeeSet = m_clsMortProdTTFees.GetTTFeeSet
End Function
Private Sub MortProdTTFees_Initialise(bIsEdit As Boolean)
    m_clsMortProdTTFees.Initialise bIsEdit
End Sub
Private Sub MortProdTTFees_InitTab()
    m_clsMortProdTTFees.InitTab
End Sub
Private Sub MortProdTTFees_SaveScreenData()
    m_clsMortProdTTFees.SaveScreenData
End Sub
Private Sub MortProdTTFees_Selected()
    m_clsMortProdTTFees.Selected
End Sub
Private Sub MortProdTTFees_SetTTFeeSet(sFeeSet As String)
    m_clsMortProdTTFees.SetTTFeeSet sFeeSet
End Sub
Private Sub MortProdTTFees_SetProductKeys(colValues As Collection)
    m_clsMortProdTTFees.SetProductKeys colValues
End Sub
Private Sub MortProdTTFees_SetScreenFields()
    m_clsMortProdTTFees.SetScreenFields
End Sub
Private Function MortProdTTFees_ValidateScreenData() As Boolean
    MortProdTTFees_ValidateScreenData = m_clsMortProdTTFees.ValidateScreenData
End Function



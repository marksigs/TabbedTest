VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "MortProdSwitchFeesCS"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'*******************************************************************************************************
' Class         : MortProdSwitchFeesCS
' Description   : Client specific version of the MortProdSwitchFeesCS class. This class will be called by
'                 a form or controlling class and any client specific code can be put in here for
'                 each client. Each method can a) override the base code, b) delegate to the base
'                 code or c) overried and delegate.
' Change history
' Prog      Date        Description
'   MC      02/06/2004  BMIDS763    BMIDS763 - CC075 - FEESETS New Class refs added
'*******************************************************************************************************
Option Explicit

' Base class
Implements MortProdSwitchFees

' Private data
Private m_clsMortProdSwitchFees As MortProdSwitchFees
Private Sub Class_Initialize()
    Set m_clsMortProdSwitchFees = New MortProdSwitchFees
End Sub
Private Sub MortProdSwitchFees_DeSelect()
    m_clsMortProdSwitchFees.DeSelect
End Sub
Private Function MortProdSwitchFees_GetProductSwitchFeeSet() As Variant
    MortProdSwitchFees_GetProductSwitchFeeSet = m_clsMortProdSwitchFees.GetProductSwitchFeeSet
End Function
Private Sub MortProdSwitchFees_Initialise(bIsEdit As Boolean)
    m_clsMortProdSwitchFees.Initialise bIsEdit
End Sub
Private Sub MortProdSwitchFees_InitTab()
    m_clsMortProdSwitchFees.InitTab
End Sub
Private Sub MortProdSwitchFees_SaveScreenData()
    m_clsMortProdSwitchFees.SaveScreenData
End Sub
Private Sub MortProdSwitchFees_Selected()
    m_clsMortProdSwitchFees.Selected
End Sub
Private Sub MortProdSwitchFees_SetProductSwitchFeeSet(sFeeSet As String)
    m_clsMortProdSwitchFees.SetProductSwitchFeeSet sFeeSet
End Sub
Private Sub MortProdSwitchFees_SetProductKeys(colValues As Collection)
    m_clsMortProdSwitchFees.SetProductKeys colValues
End Sub
Private Sub MortProdSwitchFees_SetScreenFields()
    m_clsMortProdSwitchFees.SetScreenFields
End Sub
Private Function MortProdSwitchFees_ValidateScreenData() As Boolean
    MortProdSwitchFees_ValidateScreenData = m_clsMortProdSwitchFees.ValidateScreenData
End Function





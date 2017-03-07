VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "MortProdTranEquFeeCS"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' Class         : MortProdTranEquFeeCS
' Description   : Client specific version of the MortProdTranEquFee class. This class will be called by
'                 a form or controlling class and any client specific code can be put in here for
'                 each client. Each method can a) override the base code, b) delegate to the base
'                 code or c) overried and delegate.
' Change history
' Prog      Date        Description
' TW        11/12/2006  EP2_20 - Created
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Option Explicit

' Base class
Implements MortProdTranEquFee

' Private data
Private m_clsMortProdTranEquFee As MortProdTranEquFee
Private Sub Class_Initialize()
    Set m_clsMortProdTranEquFee = New MortProdTranEquFee
End Sub
Private Sub MortProdTranEquFee_DeSelect()
    m_clsMortProdTranEquFee.DeSelect
End Sub
Private Function MortProdTranEquFee_GetTransferOfEquityFeeSet() As String
    MortProdTranEquFee_GetTransferOfEquityFeeSet = m_clsMortProdTranEquFee.GetTransferOfEquityFeeSet
End Function
Private Sub MortProdTranEquFee_Initialise(bIsEdit As Boolean)
        m_clsMortProdTranEquFee.Initialise bIsEdit
End Sub
Private Sub MortProdTranEquFee_InitTab()
    m_clsMortProdTranEquFee.InitTab
End Sub
Private Sub MortProdTranEquFee_SaveScreenData()
    m_clsMortProdTranEquFee.SaveScreenData
End Sub
Private Sub MortProdTranEquFee_Selected()
    m_clsMortProdTranEquFee.Selected
End Sub
Private Sub MortProdTranEquFee_SetFeeSet(sFeeSet As String)
    m_clsMortProdTranEquFee.SetFeeSet sFeeSet
End Sub
Private Sub MortProdTranEquFee_SetProductKeys(colValues As Collection)
    m_clsMortProdTranEquFee.SetProductKeys colValues
End Sub
Private Sub MortProdTranEquFee_SetScreenFields()
    m_clsMortProdTranEquFee.SetScreenFields
End Sub
Private Function MortProdTranEquFee_ValidateScreenData() As Boolean
    MortProdTranEquFee_ValidateScreenData = m_clsMortProdTranEquFee.ValidateScreenData
End Function


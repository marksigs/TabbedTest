VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "MortProdAddBorrFeeCS"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' Class         : MortProdAddBorrFeeCS
' Description   : Client specific version of the MortProdAddBorrFee class. This class will be called by
'                 a form or controlling class and any client specific code can be put in here for
'                 each client. Each method can a) override the base code, b) delegate to the base
'                 code or c) overried and delegate.
' Change history
' Prog      Date        Description
' TW        09/10/2006  EP2_7 Created
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Option Explicit

' Base class
Implements MortProdAddBorrFee

' Private data
Private m_clsMortProdAddBorrFee As MortProdAddBorrFee
Private Sub Class_Initialize()
    Set m_clsMortProdAddBorrFee = New MortProdAddBorrFee
End Sub
Private Sub MortProdAddBorrFee_DeSelect()
    m_clsMortProdAddBorrFee.DeSelect
End Sub
Private Function MortProdAddBorrFee_GetAdditionalBorrowingFeeSet() As String
    MortProdAddBorrFee_GetAdditionalBorrowingFeeSet = m_clsMortProdAddBorrFee.GetAdditionalBorrowingFeeSet
End Function
Private Sub MortProdAddBorrFee_Initialise(bIsEdit As Boolean)
        m_clsMortProdAddBorrFee.Initialise bIsEdit
End Sub
Private Sub MortProdAddBorrFee_InitTab()
    m_clsMortProdAddBorrFee.InitTab
End Sub
Private Sub MortProdAddBorrFee_SaveScreenData()
    m_clsMortProdAddBorrFee.SaveScreenData
End Sub
Private Sub MortProdAddBorrFee_Selected()
    m_clsMortProdAddBorrFee.Selected
End Sub
Private Sub MortProdAddBorrFee_SetFeeSet(sFeeSet As String)
    m_clsMortProdAddBorrFee.SetFeeSet sFeeSet
End Sub
Private Sub MortProdAddBorrFee_SetProductKeys(colValues As Collection)
    m_clsMortProdAddBorrFee.SetProductKeys colValues
End Sub
Private Sub MortProdAddBorrFee_SetScreenFields()
    m_clsMortProdAddBorrFee.SetScreenFields
End Sub
Private Function MortProdAddBorrFee_ValidateScreenData() As Boolean
    MortProdAddBorrFee_ValidateScreenData = m_clsMortProdAddBorrFee.ValidateScreenData
End Function

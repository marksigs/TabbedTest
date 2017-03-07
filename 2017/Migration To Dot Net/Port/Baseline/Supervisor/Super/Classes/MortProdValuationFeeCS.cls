VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "MortProdValuationFeeCS"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' Class         : MortProdValuationFeeCS
' Description   : Client specific version of the MortProdValuationFee class. This class will be called by
'                 a form or controlling class and any client specific code can be put in here for
'                 each client. Each method can a) override the base code, b) delegate to the base
'                 code or c) overried and delegate.
' Change history
' Prog      Date        Description
' DJP       10/12/01    SYS2831 Support client variants
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Option Explicit

' Base class
Implements MortProdValuationFee

' Private data
Private m_clsMortProdValuationFee As MortProdValuationFee
Private Sub Class_Initialize()
    Set m_clsMortProdValuationFee = New MortProdValuationFee
End Sub
Private Sub MortProdValuationFee_DeSelect()
    m_clsMortProdValuationFee.DeSelect
End Sub
Private Function MortProdValuationFee_GetValuationFeeSet() As String
    MortProdValuationFee_GetValuationFeeSet = m_clsMortProdValuationFee.GetValuationFeeSet
End Function
Private Sub MortProdValuationFee_Initialise(bIsEdit As Boolean)
        m_clsMortProdValuationFee.Initialise bIsEdit
End Sub
Private Sub MortProdValuationFee_InitTab()
    m_clsMortProdValuationFee.InitTab
End Sub
Private Sub MortProdValuationFee_SaveScreenData()
    m_clsMortProdValuationFee.SaveScreenData
End Sub
Private Sub MortProdValuationFee_Selected()
    m_clsMortProdValuationFee.Selected
End Sub
Private Sub MortProdValuationFee_SetFeeSet(sFeeSet As String)
    m_clsMortProdValuationFee.SetFeeSet sFeeSet
End Sub
Private Sub MortProdValuationFee_SetProductKeys(colValues As Collection)
    m_clsMortProdValuationFee.SetProductKeys colValues
End Sub
Private Sub MortProdValuationFee_SetScreenFields()
    m_clsMortProdValuationFee.SetScreenFields
End Sub
Private Function MortProdValuationFee_ValidateScreenData() As Boolean
    MortProdValuationFee_ValidateScreenData = m_clsMortProdValuationFee.ValidateScreenData
End Function

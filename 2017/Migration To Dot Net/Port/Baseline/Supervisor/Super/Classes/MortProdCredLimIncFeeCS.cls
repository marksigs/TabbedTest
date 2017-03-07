VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "MortProdCredLimIncFeeCS"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' Class         : MortProdCredLimIncFeeCS
' Description   : Client specific version of the MortProdCredLimIncFee class. This class will be called by
'                 a form or controlling class and any client specific code can be put in here for
'                 each client. Each method can a) override the base code, b) delegate to the base
'                 code or c) overried and delegate.
' Change history
' Prog      Date        Description
' TW        09/10/2006  EP2_7 Created
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Option Explicit

' Base class
Implements MortProdCredLimIncFee

' Private data
Private m_clsMortProdCredLimIncFee As MortProdCredLimIncFee
Private Sub Class_Initialize()
    Set m_clsMortProdCredLimIncFee = New MortProdCredLimIncFee
End Sub
Private Sub MortProdCredLimIncFee_DeSelect()
    m_clsMortProdCredLimIncFee.DeSelect
End Sub

Private Function MortProdCredLimIncFee_GetFeeSet() As String
    MortProdCredLimIncFee_GetFeeSet = m_clsMortProdCredLimIncFee.GetFeeSet
End Function


Private Sub MortProdCredLimIncFee_Initialise(bIsEdit As Boolean)
        m_clsMortProdCredLimIncFee.Initialise bIsEdit
End Sub
Private Sub MortProdCredLimIncFee_InitTab()
    m_clsMortProdCredLimIncFee.InitTab
End Sub
Private Sub MortProdCredLimIncFee_SaveScreenData()
    m_clsMortProdCredLimIncFee.SaveScreenData
End Sub
Private Sub MortProdCredLimIncFee_Selected()
    m_clsMortProdCredLimIncFee.Selected
End Sub
Private Sub MortProdCredLimIncFee_SetFeeSet(sFeeSet As String)
    m_clsMortProdCredLimIncFee.SetFeeSet sFeeSet
End Sub
Private Sub MortProdCredLimIncFee_SetProductKeys(colValues As Collection)
    m_clsMortProdCredLimIncFee.SetProductKeys colValues
End Sub
Private Sub MortProdCredLimIncFee_SetScreenFields()
    m_clsMortProdCredLimIncFee.SetScreenFields
End Sub
Private Function MortProdCredLimIncFee_ValidateScreenData() As Boolean
    MortProdCredLimIncFee_ValidateScreenData = m_clsMortProdCredLimIncFee.ValidateScreenData
End Function

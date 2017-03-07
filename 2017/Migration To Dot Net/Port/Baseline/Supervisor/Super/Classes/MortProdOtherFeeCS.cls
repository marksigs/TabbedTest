VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "MortProdOtherFeeCS"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' Class         : MortProdOtherFeesCS
' Description   : Client specific version of the MortProdOtherFees class. This class will be called by
'                 a form or controlling class and any client specific code can be put in here for
'                 each client. Each method can a) override the base code, b) delegate to the base
'                 code or c) overried and delegate.
' Change history
' Prog      Date        Description
' DJP       10/12/01    SYS2831 Support client variants
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''

Option Explicit

' Base class
Implements MortProdOtherFee

' Private data
Private m_clsMortProdOtherFee As MortProdOtherFee

Private Sub Class_Initialize()
    On Error GoTo Failed
    Set m_clsMortProdOtherFee = New MortProdOtherFee
    Exit Sub
Failed:
    g_clsErrorHandling.RaiseError Err.Number, Err.DESCRIPTION
End Sub

Private Sub MortProdOtherFee_DoUpdates()
    m_clsMortProdOtherFee.DoUpdates
End Sub

Private Sub MortProdOtherFee_Initialise(bIsEdit As Boolean)
    m_clsMortProdOtherFee.Initialise bIsEdit
End Sub

Private Sub MortProdOtherFee_InitTab()
    m_clsMortProdOtherFee.InitTab
End Sub

Private Sub MortProdOtherFee_SaveScreenData()
    m_clsMortProdOtherFee.SaveScreenData
End Sub

Private Function MortProdOtherFee_SetLender(vOrgID As Variant) As Boolean
    MortProdOtherFee_SetLender = m_clsMortProdOtherFee.SetLender(vOrgID)
End Function

Private Sub MortProdOtherFee_SetProductKeys(colKeyValues As Collection)
    m_clsMortProdOtherFee.SetProductKeys colKeyValues
End Sub

Private Sub MortProdOtherFee_SetScreenFields()
    m_clsMortProdOtherFee.SetScreenFields
End Sub

Private Function MortProdOtherFee_ValidateScreenData() As Boolean
    MortProdOtherFee_ValidateScreenData = m_clsMortProdOtherFee.ValidateScreenData
End Function


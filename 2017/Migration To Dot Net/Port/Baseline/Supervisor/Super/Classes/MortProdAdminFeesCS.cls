VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "MortProdAdminFeesCS"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' Class         : MortProdAdminFeesCS
' Description   : Client specific version of the MortProdAdminFees class. This class will be called by
'                 a form or controlling class and any client specific code can be put in here for
'                 each client. Each method can a) override the base code, b) delegate to the base
'                 code or c) overried and delegate.
' Change history
' Prog      Date        Description
' DJP       22/10/01    Created.
' DJP       10/12/01    SYS2831 Support client variants
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Option Explicit

' Base class
Implements MortProdAdminFees

' Private data
Private m_clsMortProdAdminFees As MortProdAdminFees
Private Sub Class_Initialize()
    Set m_clsMortProdAdminFees = New MortProdAdminFees
End Sub
Private Sub MortProdAdminFees_DeSelect()
    m_clsMortProdAdminFees.DeSelect
End Sub
Private Function MortProdAdminFees_GetAdminFeeSet() As Variant
    MortProdAdminFees_GetAdminFeeSet = m_clsMortProdAdminFees.GetAdminFeeSet
End Function
Private Sub MortProdAdminFees_Initialise(bIsEdit As Boolean)
    m_clsMortProdAdminFees.Initialise bIsEdit
End Sub
Private Sub MortProdAdminFees_InitTab()
    m_clsMortProdAdminFees.InitTab
End Sub
Private Sub MortProdAdminFees_SaveScreenData()
    m_clsMortProdAdminFees.SaveScreenData
End Sub
Private Sub MortProdAdminFees_Selected()
    m_clsMortProdAdminFees.Selected
End Sub
Private Sub MortProdAdminFees_SetAdminFeeSet(sFeeSet As String)
    m_clsMortProdAdminFees.SetAdminFeeSet sFeeSet
End Sub
Private Sub MortProdAdminFees_SetProductKeys(colValues As Collection)
    m_clsMortProdAdminFees.SetProductKeys colValues
End Sub
Private Sub MortProdAdminFees_SetScreenFields()
    m_clsMortProdAdminFees.SetScreenFields
End Sub
Private Function MortProdAdminFees_ValidateScreenData() As Boolean
    MortProdAdminFees_ValidateScreenData = m_clsMortProdAdminFees.ValidateScreenData
End Function

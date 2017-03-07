VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "MortProdIncentiveCS"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' Class         : MortProdIncentiveCS
' Description   : Client specific version of the MortProdIncentive class. This class will be called by
'                 a form or controlling class and any client specific code can be put in here for
'                 each client. Each method can a) override the base code, b) delegate to the base
'                 code or c) overried and delegate.
' Change history
' Prog      Date        Description
' DJP       10/12/01    SYS2831 Support client variants
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''

Option Explicit
' Base class
Implements MortProdIncentive

' Private data
Private m_clsIncentive As MortProdIncentive
Private Sub Class_Initialize()
    Set m_clsIncentive = New MortProdIncentive
End Sub
Private Sub MortProdIncentive_Add()
    m_clsIncentive.Add
End Sub
Private Sub MortProdIncentive_Delete()
    m_clsIncentive.Delete
End Sub
Private Sub MortProdIncentive_Edit()
    m_clsIncentive.Edit
End Sub
Private Sub MortProdIncentive_Initialise(bIsEdit As Boolean)
    m_clsIncentive.Initialise bIsEdit
End Sub
Private Sub MortProdIncentive_InitTab()
    m_clsIncentive.InitTab
End Sub
Private Sub MortProdIncentive_SaveScreenData()
    m_clsIncentive.SaveScreenData
End Sub
Private Sub MortProdIncentive_SetProductKeys(colValues As Collection)
    m_clsIncentive.SetProductKeys colValues
End Sub
Private Function MortProdIncentive_SetScreenFields() As Variant
    MortProdIncentive_SetScreenFields = m_clsIncentive.SetScreenFields
End Function
Private Function MortProdIncentive_ValidateScreenData() As Boolean
    MortProdIncentive_ValidateScreenData = m_clsIncentive.ValidateScreenData
End Function

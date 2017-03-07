VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "MortProdTypeOfBuyerCS"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' Class         : MortProdTypeOfBuyerCS
' Description   : Client specific version of the MortProdTypeOfBuyer class. This class will be called by
'                 a form or controlling class and any client specific code can be put in here for
'                 each client. Each method can a) override the base code, b) delegate to the base
'                 code or c) overried and delegate.
' Change history
' Prog      Date        Description
' DJP       10/12/01    SYS2831 Support client variants
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Option Explicit

' Base class
Implements MortProdTypeOfBuyer

' Private data
Private m_clsMortProdTypeOfBuyer As MortProdTypeOfBuyer
Private Sub Class_Initialize()
    Set m_clsMortProdTypeOfBuyer = New MortProdTypeOfBuyer
End Sub
Private Sub MortProdTypeOfBuyer_DoUpdates()
    m_clsMortProdTypeOfBuyer.DoUpdates
End Sub
Private Sub MortProdTypeOfBuyer_Initialise(bIsEdit As Boolean)
    m_clsMortProdTypeOfBuyer.Initialise bIsEdit
End Sub
Private Sub MortProdTypeOfBuyer_InitTab()
    m_clsMortProdTypeOfBuyer.InitTab
End Sub
Private Sub MortProdTypeOfBuyer_SaveScreenData()
    m_clsMortProdTypeOfBuyer.SaveScreenData
End Sub
Private Sub MortProdTypeOfBuyer_SetProductKeys(colValues As Collection)
    m_clsMortProdTypeOfBuyer.SetProductKeys colValues
End Sub
Private Sub MortProdTypeOfBuyer_SetScreenFields()
    m_clsMortProdTypeOfBuyer.SetScreenFields
End Sub
Private Function MortProdTypeOfBuyer_ValidateScreenData() As Variant
    MortProdTypeOfBuyer_ValidateScreenData = m_clsMortProdTypeOfBuyer.ValidateScreenData
End Function

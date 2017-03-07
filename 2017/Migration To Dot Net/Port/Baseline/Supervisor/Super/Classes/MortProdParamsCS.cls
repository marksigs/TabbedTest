VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "MortProdParamsCS"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' Class         : MortProdParamsCS
' Description   : Client specific version of the MortProdParams class. This class will be called by
'                 a form or controlling class and any client specific code can be put in here for
'                 each client. Each method can a) override the base code, b) delegate to the base
'                 code or c) overried and delegate.
' Change history
' Prog      Date        Description
' DJP       10/12/01    SYS2831 Support client variants
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Option Explicit

' Base class
Implements AdditionalMortProdParams

' Private data
Private m_clsAdditionalParams As AdditionalMortProdParams
Private Sub AdditionalMortProdParams_InitTab()
    m_clsAdditionalParams.InitTab
End Sub
Private Sub Class_Initialize()
    Set m_clsAdditionalParams = New AdditionalMortProdParams
End Sub
Private Sub AdditionalMortProdParams_DoUpdates()
    m_clsAdditionalParams.DoUpdates
End Sub
Private Sub AdditionalMortProdParams_Initialise(bIsEdit As Boolean)
    m_clsAdditionalParams.Initialise bIsEdit
End Sub
Private Sub AdditionalMortProdParams_SaveScreenData()
    m_clsAdditionalParams.SaveScreenData
End Sub
Private Sub AdditionalMortProdParams_SetProductKeys(colValues As Collection)
    m_clsAdditionalParams.SetProductKeys colValues
End Sub
Private Sub AdditionalMortProdParams_SetScreenFields()
    m_clsAdditionalParams.SetScreenFields
End Sub
Private Function AdditionalMortProdParams_ValidateScreenData() As Boolean
    AdditionalMortProdParams_ValidateScreenData = m_clsAdditionalParams.ValidateScreenData
End Function

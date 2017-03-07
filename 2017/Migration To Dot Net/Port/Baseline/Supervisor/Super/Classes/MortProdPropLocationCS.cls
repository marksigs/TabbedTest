VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "MortProdPropLocationCS"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' Class         : MortProdPropLocationCS
' Description   : Client specific version of the MortProdPropLocation class. This class will be called by
'                 a form or controlling class and any client specific code can be put in here for
'                 each client. Each method can a) override the base code, b) delegate to the base
'                 code or c) overried and delegate.
' Change history
' Prog      Date        Description
' DJP       10/12/01    SYS2831 Support client variants
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''

Option Explicit

' Base class
Implements MortProdPropLocation

' Private data
Private m_clsMortProdPropLocation As MortProdPropLocation
Private Sub Class_Initialize()
    On Error GoTo Failed
    
    Set m_clsMortProdPropLocation = New MortProdPropLocation
    
    Exit Sub
Failed:
    g_clsErrorHandling.RaiseError Err.Number, Err.DESCRIPTION
End Sub
Private Sub MortProdPropLocation_DoUpdates()
    m_clsMortProdPropLocation.DoUpdates
End Sub
Private Sub MortProdPropLocation_Initialise(bIsEdit As Boolean)
    m_clsMortProdPropLocation.Initialise bIsEdit
End Sub
Private Sub MortProdPropLocation_InitTab()
    m_clsMortProdPropLocation.InitTab
End Sub
Private Sub MortProdPropLocation_SaveScreenData()
    m_clsMortProdPropLocation.SaveScreenData
End Sub
Private Sub MortProdPropLocation_SetProductKeys(colValues As Collection)
    m_clsMortProdPropLocation.SetProductKeys colValues
End Sub
Private Sub MortProdPropLocation_SetScreenFields()
    m_clsMortProdPropLocation.SetScreenFields
End Sub
Private Function MortProdPropLocation_ValidateScreenData() As Boolean
    MortProdPropLocation_ValidateScreenData = m_clsMortProdPropLocation.ValidateScreenData
End Function

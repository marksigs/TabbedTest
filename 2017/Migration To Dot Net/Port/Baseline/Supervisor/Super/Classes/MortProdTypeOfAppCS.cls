VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "MortProdTypeOfAppCS"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' Class         : MortProdTypeOfAppCS
' Description   : Client specific version of the MortProdTypeOfAppElig class. This class will be called by
'                 a form or controlling class and any client specific code can be put in here for
'                 each client. Each method can a) override the base code, b) delegate to the base
'                 code or c) overried and delegate.
' Change history
' Prog      Date        Description
' DJP       10/12/01    SYS2831 Support client variants
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''

Option Explicit

' Base class
Implements MortProdTypeofAppElig

' Private data
Private m_clsMortProdTypeOfApp As MortProdTypeofAppElig
Private Sub Class_Initialize()
    Set m_clsMortProdTypeOfApp = New MortProdTypeofAppElig
End Sub
Private Sub MortProdTypeofAppElig_DoUpdates()
    m_clsMortProdTypeOfApp.DoUpdates
End Sub
Private Sub MortProdTypeofAppElig_Initialise(bIsEdit As Boolean)
    m_clsMortProdTypeOfApp.Initialise bIsEdit
End Sub
Private Sub MortProdTypeofAppElig_InitTab()
    m_clsMortProdTypeOfApp.InitTab
End Sub
Private Sub MortProdTypeofAppElig_SaveScreenData()
    m_clsMortProdTypeOfApp.SaveScreenData
End Sub
Private Sub MortProdTypeofAppElig_SetProductKeys(colValues As Collection)
    m_clsMortProdTypeOfApp.SetProductKeys colValues
End Sub
Private Function MortProdTypeofAppElig_SetScreenFields() As Boolean
    MortProdTypeofAppElig_SetScreenFields = m_clsMortProdTypeOfApp.SetScreenFields
End Function
Private Function MortProdTypeofAppElig_ValidateScreenData() As Boolean
    MortProdTypeofAppElig_ValidateScreenData = m_clsMortProdTypeOfApp.ValidateScreenData
End Function

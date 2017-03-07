VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "MortProdChannelEligCS"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' Class         : MortProdChannelEligCS
' Description   : Client specific version of the MortProdChannelElig class. This class will be called by
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
Implements MortProdChannelElig

' Private data
Private m_clsMortProdChannelElig As MortProdChannelElig
Private Sub Class_Initialize()
    Set m_clsMortProdChannelElig = New MortProdChannelElig
End Sub
Private Sub MortProdChannelElig_Initialise(bIsEdit As Boolean)
    m_clsMortProdChannelElig.Initialise bIsEdit
End Sub
Private Sub MortProdChannelElig_InitTab()
    m_clsMortProdChannelElig.InitTab
End Sub
Private Sub MortProdChannelElig_SaveScreenData()
    m_clsMortProdChannelElig.SaveScreenData
End Sub
Private Sub MortProdChannelElig_SetProductKeys(colValues As Collection)
    m_clsMortProdChannelElig.SetProductKeys colValues
End Sub
Private Sub MortProdChannelElig_SetScreenFields()
    m_clsMortProdChannelElig.SetScreenFields
End Sub
Private Function MortProdChannelElig_ValidateScreenData() As Variant
    MortProdChannelElig_ValidateScreenData = m_clsMortProdChannelElig.ValidateScreenData
End Function

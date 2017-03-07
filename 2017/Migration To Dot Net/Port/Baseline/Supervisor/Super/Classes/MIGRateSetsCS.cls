VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "MIGRateSetsCS"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' Class         : MIGRatesCS
' Description   : Client specific version of the MIGRates class. This class will be called by
'                 a form or controlling class and any client specific code can be put in here for
'                 each client. Each method can a) override the base code, b) delegate to the base
'                 code or c) overried and delegate.
' Change history
' Prog      Date        Description
' DJP       22/10/01    Created.
' DJP       20/11/01    SYS2831/SYS2912 Support client variants & SQL Server locking problem.
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Option Explicit
Implements MIGRateSets

' Private data
Private m_clsMIGRates As MIGRateSets
Private Sub Class_Initialize()
    Set m_clsMIGRates = New MIGRateSets
End Sub
Private Sub MIGRateSets_Add()
    m_clsMIGRates.Add
End Sub
Private Sub MIGRateSets_Delete()
    m_clsMIGRates.Delete
End Sub
Private Sub MIGRateSets_DoUpdates()
    m_clsMIGRates.DoUpdates
End Sub
Private Sub MIGRateSets_Edit()
    m_clsMIGRates.Edit
End Sub
Private Function MIGRateSets_GetRateSet() As String
    MIGRateSets_GetRateSet = m_clsMIGRates.GetRateSet
End Function
Private Sub MIGRateSets_Initialise(bIsEdit As Boolean)
    m_clsMIGRates.Initialise bIsEdit
End Sub
Private Sub MIGRateSets_InitTab()
    m_clsMIGRates.InitTab
End Sub
Private Sub MIGRateSets_SaveScreenData()
    m_clsMIGRates.SaveScreenData
End Sub
Private Sub MIGRateSets_SetButtonsState()
    m_clsMIGRates.SetButtonsState
End Sub
Private Sub MIGRateSets_SetMIGRateSet(vRateSet As Variant)
    m_clsMIGRates.SetMIGRateSet vRateSet
End Sub
Private Sub MIGRateSets_SetScreenFields()
    m_clsMIGRates.SetScreenFields
End Sub
Private Function MIGRateSets_ValidateScreenData() As Boolean
    MIGRateSets_ValidateScreenData = m_clsMIGRates.ValidateScreenData
End Function

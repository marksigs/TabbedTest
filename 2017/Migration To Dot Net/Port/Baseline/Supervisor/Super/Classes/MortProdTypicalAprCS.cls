VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "MortProdTypicalAprCS"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' Class         : MortProdTypicalAPRCS
' Description   : Client specific version of the MortProdTypicalAPR class. This class will be called by
'                 a form or controlling class and any client specific code can be put in here for
'                 each client. Each method can a) override the base code, b) delegate to the base
'                 code or c) overried and delegate.
' Change history
' Prog      Date        Description
' DJP       10/12/01    SYS2831 Support client variants
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Option Explicit

' Base class
Implements MortProdTypicalAPR

' Private data
Private m_clsMortProdTypicalAPR As MortProdTypicalAPR
Private Sub Class_Initialize()
    On Error GoTo Failed
    
    Set m_clsMortProdTypicalAPR = New MortProdTypicalAPR
    
    Exit Sub
Failed:
    g_clsErrorHandling.RaiseError Err.Number, Err.DESCRIPTION
End Sub
Private Sub MortProdTypicalAPR_DoUpdates()
    m_clsMortProdTypicalAPR.DoUpdates
End Sub
Private Sub MortProdTypicalAPR_Initialise(Optional bIsEdit As Boolean)
    m_clsMortProdTypicalAPR.Initialise bIsEdit
End Sub
Private Sub MortProdTypicalAPR_InitTab()
    m_clsMortProdTypicalAPR.InitTab
End Sub
Private Sub MortProdTypicalAPR_SaveScreenData()
    m_clsMortProdTypicalAPR.SaveScreenData
End Sub
Private Sub MortProdTypicalAPR_SetProductKeys(colValues As Collection)
    m_clsMortProdTypicalAPR.SetProductKeys colValues
End Sub
Private Sub MortProdTypicalAPR_SetScreenFields()
    m_clsMortProdTypicalAPR.SetScreenFields
End Sub
Private Function MortProdTypicalAPR_ValidateScreenData() As Boolean
    MortProdTypicalAPR_ValidateScreenData = m_clsMortProdTypicalAPR.ValidateScreenData
End Function

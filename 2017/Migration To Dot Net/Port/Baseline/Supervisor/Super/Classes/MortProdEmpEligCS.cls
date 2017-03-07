VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "MortProdEmpEligCS"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' Class         : MortProdEmpEligCS
' Description   : Client specific version of the MortProdEmpElig class. This class will be called by
'                 a form or controlling class and any client specific code can be put in here for
'                 each client. Each method can a) override the base code, b) delegate to the base
'                 code or c) overried and delegate.
' Change history
' Prog      Date        Description
' DJP       10/12/01    SYS2831 Support client variants
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''

Option Explicit

'Base Class
Implements MortProdEmpElig

' Private data
Private m_clsMortProdEmpElig As MortProdEmpElig
Private Sub Class_Initialize()
    Set m_clsMortProdEmpElig = New MortProdEmpElig
End Sub
Private Sub MortProdEmpElig_Initialise(bIsEdit As Boolean)
    m_clsMortProdEmpElig.Initialise bIsEdit
End Sub
Private Sub MortProdEmpElig_InitTab()
    m_clsMortProdEmpElig.InitTab
End Sub
Private Sub MortProdEmpElig_SaveScreenData()
    m_clsMortProdEmpElig.SaveScreenData
End Sub
Private Sub MortProdEmpElig_SetProductKeys(colValues As Collection)
    m_clsMortProdEmpElig.SetProductKeys colValues
End Sub
Private Sub MortProdEmpElig_SetScreenFields()
    m_clsMortProdEmpElig.SetScreenFields
End Sub
Private Function MortProdEmpElig_ValidateScreenData() As Boolean
    MortProdEmpElig_ValidateScreenData = m_clsMortProdEmpElig.ValidateScreenData
End Function

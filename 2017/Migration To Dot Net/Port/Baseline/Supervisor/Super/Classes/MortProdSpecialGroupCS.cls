VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "MortProdSpecialGroupCS"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' Class         : MortProdSpecialGroupCS
' Description   : Client specific version of the MortProdSpecialGroup class. This class will be called by
'                 a form or controlling class and any client specific code can be put in here for
'                 each client. Each method can a) override the base code, b) delegate to the base
'                 code or c) overried and delegate.
' Change history
' Prog      Date        Description
' DJP       10/12/01    SYS2831 Support client variants
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Option Explicit

' Base class
Implements MortProdSpecialGroup

' Private data
Private m_clsMortProdSpecialGroup As MortProdSpecialGroup

Private Sub Class_Initialize()
    Set m_clsMortProdSpecialGroup = New MortProdSpecialGroup
End Sub

Private Sub MortProdSpecialGroup_DoUpdates()
    m_clsMortProdSpecialGroup.DoUpdates
End Sub

Private Sub MortProdSpecialGroup_Initialise(Optional bIsEdit As Boolean)
    m_clsMortProdSpecialGroup.Initialise bIsEdit
End Sub

Private Sub MortProdSpecialGroup_InitTab()
    m_clsMortProdSpecialGroup.InitTab
End Sub

Private Function MortProdSpecialGroup_IsEdit() As Variant
    MortProdSpecialGroup_IsEdit = m_clsMortProdSpecialGroup.IsEdit
End Function

Private Sub MortProdSpecialGroup_SaveScreenData()
    m_clsMortProdSpecialGroup.SaveScreenData
End Sub

Private Sub MortProdSpecialGroup_SetIsEdit(Optional bEdit As Boolean = True)
    m_clsMortProdSpecialGroup.SetIsEdit bEdit
End Sub

Private Sub MortProdSpecialGroup_SetProductKeys(colValues As Collection)
    m_clsMortProdSpecialGroup.SetProductKeys colValues
End Sub

Private Sub MortProdSpecialGroup_SetScreenFields()
    m_clsMortProdSpecialGroup.SetScreenFields
End Sub

Private Function MortProdSpecialGroup_ValidateScreenData() As Variant
    MortProdSpecialGroup_ValidateScreenData = m_clsMortProdSpecialGroup.ValidateScreenData
End Function


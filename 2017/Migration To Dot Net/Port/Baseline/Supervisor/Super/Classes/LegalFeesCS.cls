VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "LegalFeesCS"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' Class         : LegalFeesCS
' Description   : Client specific version of the LegalFees class. This class will be called by
'                 a form or controlling class and any client specific code can be put in here for
'                 each client. Each method can a) override the base code, b) delegate to the base
'                 code or c) overried and delegate.
' Change history
' Prog      Date        Description
' DJP       22/10/01    Created.
' DJP       20/11/01    SYS2831/SYS2912 Support client variants & SQL Server locking problem.
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Option Explicit

Implements LegalFees

' Private data
Private m_clsLegalFees As LegalFees
Private Sub Class_Initialize()
    Set m_clsLegalFees = New LegalFees
End Sub
Private Sub LegalFees_Add()
    m_clsLegalFees.Add
End Sub

Private Sub LegalFees_Delete()
    m_clsLegalFees.Delete
End Sub
Private Sub LegalFees_DoUpdates()
    m_clsLegalFees.DoUpdates
End Sub
Private Sub LegalFees_Edit()
    m_clsLegalFees.Edit
End Sub
Private Sub LegalFees_Initialise(bIsEdit As Boolean)
    m_clsLegalFees.Initialise bIsEdit
End Sub
Private Sub LegalFees_InitTab()
    m_clsLegalFees.InitTab
End Sub
Private Function LegalFees_IsEdit() As Variant
    LegalFees_IsEdit = m_clsLegalFees.IsEdit
End Function
Private Sub LegalFees_SaveScreenData()
    m_clsLegalFees.SaveScreenData
End Sub
Private Sub LegalFees_SetButtonsState()
    m_clsLegalFees.SetButtonsState
End Sub
Private Sub LegalFees_SetIsEdit(Optional bEdit As Boolean = True)
    m_clsLegalFees.SetIsEdit bEdit
End Sub
Private Sub LegalFees_SetOrganisationID(vOrgID As Variant)
    m_clsLegalFees.SetOrganisationID vOrgID
End Sub
Private Sub LegalFees_SetScreenFields()
    m_clsLegalFees.SetScreenFields
End Sub
Private Function LegalFees_ValidateScreenData() As Boolean
    LegalFees_ValidateScreenData = m_clsLegalFees.ValidateScreenData
End Function

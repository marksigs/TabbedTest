VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "OtherFeesCS"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' Class         : OtherFeesCS
' Description   : Client specific version of the OtherFees class. This class will be called by
'                 a form or controlling class and any client specific code can be put in here for
'                 each client. Each method can a) override the base code, b) delegate to the base
'                 code or c) overried and delegate.
' Change history
' Prog      Date        Description
' DJP       22/10/01    Created.
' DJP       20/11/01    SYS2831/SYS2912 Support client variants & SQL Server locking problem.
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Option Explicit
Implements OtherFees
Private m_clsOtherFees As OtherFees
Private Sub Class_Initialize()
    Set m_clsOtherFees = New OtherFees
End Sub
Private Sub OtherFees_DoUpdates()
    m_clsOtherFees.DoUpdates
End Sub
Private Sub OtherFees_Initialise(Optional bIsEdit As Boolean)
    m_clsOtherFees.Initialise bIsEdit
End Sub
Private Sub OtherFees_InitTab()
    m_clsOtherFees.InitTab
End Sub
Private Sub OtherFees_SetKeyValues(colValues As Collection)
    m_clsOtherFees.SetKeyValues colValues
End Sub
Private Sub OtherFees_SetScreenFields()
    m_clsOtherFees.SetScreenFields
End Sub
Private Sub OtherFees_ValidateScreenData()
    m_clsOtherFees.ValidateScreenData
End Sub

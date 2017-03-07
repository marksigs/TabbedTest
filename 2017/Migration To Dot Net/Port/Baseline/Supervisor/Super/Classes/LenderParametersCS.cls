VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "LenderParametersCS"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' Class         : LenderParametersCS
' Description   : Client specific version of the LenderParameters class. This class will be called by
'                 a form or controlling class and any client specific code can be put in here for
'                 each client. Each method can a) override the base code, b) delegate to the base
'                 code or c) overried and delegate.
' Change history
' Prog      Date        Description
' DJP       22/10/01    Created.
' DJP       20/11/01    SYS2831/SYS2912 Support client variants & SQL Server locking problem.
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Option Explicit
' Base class
Implements AdditionalLenderParameters

' Private data
Private m_clsAdditionalParams As AdditionalLenderParameters
Private Sub AdditionalLenderParameters_SetKeyValues(colValues As Collection)
    m_clsAdditionalParams.SetKeyValues colValues
End Sub
Private Sub Class_Initialize()
    Set m_clsAdditionalParams = New AdditionalLenderParameters
End Sub
Private Sub AdditionalLenderParameters_InitTab()
    m_clsAdditionalParams.InitTab
End Sub
Private Function AdditionalLenderParameters_IsEdit() As Boolean
    AdditionalLenderParameters_IsEdit = m_clsAdditionalParams.IsEdit
End Function
Private Sub AdditionalLenderParameters_SaveScreenData()
    m_clsAdditionalParams.SaveScreenData
End Sub
Private Sub AdditionalLenderParameters_SetIsEdit(Optional bEdit As Boolean = True)
    m_clsAdditionalParams.SetIsEdit bEdit
End Sub
Private Sub AdditionalLenderParameters_SetScreenFields()
    m_clsAdditionalParams.SetScreenFields
End Sub
Private Function AdditionalLenderParameters_ValidateScreenData() As Boolean
    AdditionalLenderParameters_ValidateScreenData = m_clsAdditionalParams.ValidateScreenData
End Function
Private Sub AdditionalLenderParameters_DoUpdates()
    m_clsAdditionalParams.DoUpdates
End Sub
Private Sub AdditionalLenderParameters_Initialise(bIsEdit As Boolean)
    m_clsAdditionalParams.Initialise bIsEdit
End Sub

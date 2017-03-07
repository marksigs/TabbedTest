VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "ObjectDependenciesCS"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' Class         : ObjectDependenciesCS
' Description   : Provides client specific functionality for dependencies within the Promotion
'                 functionality.
' Change history
' Prog      Date        Description
' DJP       17/12/01    SYS4121 Created.
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''

Option Explicit
' Private data
Dim m_clsPromote As SupervisorObjectCopy
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' Function      :   CheckDependencies
' Description   :   Client specific version of the Core method of the same name. Check for any dependencies
'                   that exist and Promote them.
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' BMIDS Change history
' Prog      Date        Description
' GD        27/05/02    AQR : BMIDS00016 ; Supervisor Promotions

Public Function CheckDependencies(clsUpdateDetails As SupervisorUpdateDetails, colUpdateDetails As Collection) As Boolean
    On Error GoTo Failed
    Dim sObject As String
    Dim bDependencyFound As Boolean
    
    bDependencyFound = False
    
    sObject = clsUpdateDetails.GetObjectName()
    
'    Select Case sObject
'        Case REDEM_FEE_SETS
'            bDependencyFound = CheckRedemptionFeeSet(clsUpdateDetails, colUpdateDetails)
'
'        Case MP_MIG_RATE_SETS
'            bDependencyFound = CheckMPMIGRateSet(clsUpdateDetails, colUpdateDetails)
'
'        Case INCOME_MULTIPLE
'            bDependencyFound = CheckIncomeMultipleSet(clsUpdateDetails, colUpdateDetails)
'
'    End Select
    
    CheckDependencies = bDependencyFound
    
    Exit Function
Failed:
    g_clsErrorHandling.RaiseError Err.Number, Err.DESCRIPTION
End Function


Private Sub Class_Initialize()
    Set m_clsPromote = New SupervisorObjectCopy
End Sub

Private Sub Class_Terminate()
    Set m_clsPromote = Nothing
End Sub
'CheckRedemptionFeeSet(clsUpdateDetails, colUpdateDetails)
'CheckMPMIGRateSet(clsUpdateDetails, colUpdateDetails)
'CheckIncomeMultipleSet(clsUpdateDetails, colUpdateDetails)
'Private Function CheckBaseRate(ByRef clsUpdateDetails As SupervisorUpdateDetails, ByRef colUpdateDetails As Collection) As Boolean
'
'    Dim bDependencyFound As Boolean
'    Dim colObjectList As Collection
'    Dim colMatchValues As Collection
'
'    'Check for combogroup dependancies.
'    Set colObjectList = New Collection
'
'    'UserAccessType.
'    Set colMatchValues = New Collection
'    colMatchValues.Add "RateType"
'    colObjectList.Add colMatchValues
'
'    'ComboGroups.
'    bDependencyFound = GetObjectDependencies(COMBOBOX_ENTRIES, colUpdateDetails, colObjectList) Or bDependencyFound
'
'    CheckBaseRate = bDependencyFound
'
'End Function

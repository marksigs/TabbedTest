VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "PromoteCS"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' Class         : PromoteCS
' Description   : Provides client specific functionality for Promotion functionality of Supervisor.
' Change history
' Prog      Date        Description
' DJP       17/12/01    SYS4121 Created.
' DJP       24/02/02    SYS2936 WP6 Tax Authorities.
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Option Explicit
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' Function      : GetObjectTableClass
' Description   : Given the object type passed in, return the Table Class associated with it. This
'                 is a client specific version of the Core method.
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''

' BMIDS Change history
' Prog      Date        Description
' GD        27/05/01    BMIDS00016 ; Supervisor Promotions

Public Function GetObjectTableClass(sObjectName As String, Optional bTableForEdit As Boolean = True) As TableAccess
    On Error GoTo Failed
    Dim clsTableAccess As TableAccess
    
    Select Case sObjectName
    
    End Select
    
    Set GetObjectTableClass = clsTableAccess
    
    Exit Function
Failed:
    g_clsErrorHandling.RaiseError Err.Number, Err.DESCRIPTION
End Function

Public Function GetObjectDescription(sTable As String) As String
'GD BMIDS00016
    On Error GoTo Failed
    Dim sDescription As String
    
    Select Case sTable
            Case "MORTGAGEPRODUCTCONDITION"
            sDescription = MORTGAGEPRODUCTCONDITION
    End Select
    
    GetObjectDescription = sDescription
    
    Exit Function
Failed:
    g_clsErrorHandling.RaiseError Err.Number, Err.DESCRIPTION
End Function
Public Function SetupListViewHeaders(sObjectName As String, lvHeaders As listViewAccess, colHeaders As Collection) As Boolean

    On Error GoTo Failed
    Dim bSuccess As Boolean
        
    bSuccess = True
'GD 27/05/02 AQR : BMIDS00016 ; Supervisor Promotions
    Select Case sObjectName
            Case REDEM_FEE_SETS
                SetRedemptionFeeSetHeaders lvHeaders, colHeaders
            
            Case MP_MIG_RATE_SETS
                SetMPMIGRateSetHeaders lvHeaders, colHeaders
                
            Case INCOME_MULTIPLE
                SetIncomeMultipleSetHeaders lvHeaders, colHeaders
                
            Case Else
            bSuccess = False
    End Select
    
    SetupListViewHeaders = bSuccess
    
    Exit Function
Failed:
    g_clsErrorHandling.RaiseError Err.Number, Err.DESCRIPTION
End Function

Public Function PromoteObject(clsPromote As SupervisorObjectCopy, clsTableAccess As TableAccess, sTargetDatabase As String) As Boolean

On Error GoTo Failed
    Dim sTable As String
    Dim bSuccess As Boolean
    bSuccess = True

    sTable = clsTableAccess.GetTable
'GD 27/05/02 AQR : BMIDS00016 ; Supervisor Promotions
    Select Case sTable
        Case "INCOMEMULTIPLESET"
                clsPromote.PromoteIncomeMultipleSet clsTableAccess, sTargetDatabase
        Case "MORTGAGEPRODUCTCONDITION"
                clsPromote.PromoteMortgageProductCondition clsTableAccess, sTargetDatabase
                
        Case Else
            bSuccess = False
    End Select

    PromoteObject = bSuccess

Exit Function
Failed:
    g_clsErrorHandling.RaiseError Err.Number, Err.DESCRIPTION
End Function

''GD 27/05/02 AQR : BMIDS00016 ; Supervisor Promotions - SetRedemptionFeeSetHeaders
Private Sub SetRedemptionFeeSetHeaders(lvHeaders As listViewAccess, colHeaders As Collection)
    On Error GoTo Failed

    lvHeaders.nWidth = 20
    lvHeaders.sName = "Set Number"

    colHeaders.Add lvHeaders

    Exit Sub
Failed:
    g_clsErrorHandling.RaiseError Err.Number, Err.DESCRIPTION
End Sub


''GD 27/05/02 AQR : BMIDS00016 ; Supervisor Promotions - SetMPMIGRateSetHeaders
Private Sub SetMPMIGRateSetHeaders(lvHeaders As listViewAccess, colHeaders As Collection)
    On Error GoTo Failed

    lvHeaders.nWidth = 20
    lvHeaders.sName = "Set Number"

    colHeaders.Add lvHeaders

    Exit Sub
Failed:
    g_clsErrorHandling.RaiseError Err.Number, Err.DESCRIPTION
End Sub
''GD 27/05/02 AQR : BMIDS00016 ; Supervisor Promotions - SetIncomeMultipleSetHeaders
Private Sub SetIncomeMultipleSetHeaders(lvHeaders As listViewAccess, colHeaders As Collection)
    On Error GoTo Failed

    lvHeaders.nWidth = 20
    lvHeaders.sName = "Income Multiplier Code"

    colHeaders.Add lvHeaders

    Exit Sub
Failed:
    g_clsErrorHandling.RaiseError Err.Number, Err.DESCRIPTION
End Sub


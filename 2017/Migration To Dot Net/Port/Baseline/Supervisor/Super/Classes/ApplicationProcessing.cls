VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "ApplicationProcessing"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' Class         : ApplicationProcessing
' Description   : Contains functions associated with applications, such as the
'                 ability to cancel an application
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Option Explicit
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' Function      : CancelApplication
' Description   : Cancel an application; a cancelled application is one that has an entry
'                 in the applicationstage table for Cancelled, which is obtained from
'                 the combogroup APPSTAGE_COMBO
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Public Sub CancelApplication(sAppNo As String)
    On Error GoTo Failed
    Dim clsAppStage As AppStageTable
    Dim sStageNo As String
    Dim sStageName As String
    Dim clsCombo As ComboValidationTable
    Dim sFunctionName As String
    Dim bStageExists As Boolean
    
    sFunctionName = "CancelApplication"
    
    ' First, find valueid and name the validation type for cancelled (C)
    Set clsCombo = New ComboValidationTable
    clsCombo.GetValueFromValidation APPSTAGE_COMBO, STAGE_CANCELLED
    
    If TableAccess(clsCombo).RecordCount() = 0 Then
        g_clsErrorHandling.RaiseError errGeneralError, sFunctionName + " Unable to locate combo details for validation type: " + STAGE_CANCELLED
    End If
    
    sStageNo = clsCombo.GetValueID()
    sStageName = clsCombo.GetValueName()
    
    Set clsAppStage = New AppStageTable
    
    ' If the stage exists already, just return
    bStageExists = clsAppStage.DoesStageExist(sAppNo, sStageNo)
    
    If Not bStageExists Then
        clsAppStage.SetStage sAppNo, sStageNo, sStageName
    End If
    
    Exit Sub
Failed:
    g_clsErrorHandling.RaiseError Err.Number, Err.DESCRIPTION
End Sub

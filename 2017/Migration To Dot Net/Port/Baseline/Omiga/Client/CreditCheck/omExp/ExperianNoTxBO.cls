VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 1  'NoTransaction
END
Attribute VB_Name = "ExperianNoTxBO"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
'Workfile:      ExperianNoTxBO.cls
'Copyright:     Copyright © 2004 Marlborough Stirling

'Description:   Experian Business Object (No Transactions). BMids specific.

'-------------------------------------------------------------------------------------------------------
Option Explicit

Implements ObjectControl
Implements IExperianNoTxBO

Private m_objContext    As ObjectContext
Private m_objErrAssist  As ErrAssist


'----------------------------------------------------------------------------------
'BMIDS Specific History:
'
'Prog   Date        AQR         Description
'
'HMA    04/08/2004  BMIDS748    Created to write Experian error to memo pad.
'
'----------------------------------------------------------------------------------
Private Function IExperianNoTxBO_CreateMemoPad(vstrRequest As String) As String

    '
    'This function calls omApp.ApplicationBo.CreateMemoPad to write the error returned by Experian
    '(during an UpgradeToFullAppCreditCheck) to the Memo Pad.
    '

    On Error GoTo CreateMemoPadVbErr
    Const cstrFunctionName As String = "IExperianNoTxBO_CreateMemoPad"
    
    Dim objAppBO As ApplicationBO
    
    Set objAppBO = m_objContext.CreateInstance(gstrAPPLICATION_COMPONENT & ".ApplicationBO")
    
    IExperianNoTxBO_CreateMemoPad = objAppBO.CreateMemoPad(vstrRequest)
    
CreateMemoPadExit:

    If Not m_objContext Is Nothing Then
        m_objContext.SetComplete
    End If

    Set objAppBO = Nothing
    
    Exit Function
    
CreateMemoPadVbErr:

    If m_objErrAssist.IsSystemError = True Then
        m_objErrAssist.LogError TypeName(Me), cstrFunctionName, Err.Number, Err.Description, vbLogEventTypeError
    End If

    Err.Raise Err.Number, Err.Source, Err.Description
    
    'go to clean-up section
    Resume CreateMemoPadExit

End Function

Private Sub ObjectControl_Activate()
    Set m_objContext = GetObjectContext()
End Sub

Private Function ObjectControl_CanBePooled() As Boolean
    ObjectControl_CanBePooled = False
End Function

Private Sub ObjectControl_Deactivate()
    Set m_objContext = Nothing
End Sub

VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "IncomeFactorCS"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' Form          : IncomeFactorCS
' Description   : Client-specific version of the IncomeFactor support class.
'                 This class inherits the interface and aggragates the base
'                 class.
'
' Change history
' Prog  Date        Description
' STB   14-May-2002 SYS4417 Added AllowableIncomeFactors.
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Option Explicit

'Inherit the niterface of the base class.
Implements IncomeFactor

'Aggregated base class.
Private m_clsIncomeFactor As IncomeFactor


'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' Function    : Class_Initialize
' Description : Create the aggregated base class.
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub Class_Initialize()
    Set m_clsIncomeFactor = New IncomeFactor
End Sub


'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' Function    : Class_Terminate
' Description : Release the base class reference.
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub Class_Terminate()
    Set m_clsIncomeFactor = Nothing
End Sub


'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' Function    : IncomeFactor_GetReturnCode
' Description : Delegate to the base class.
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Function IncomeFactor_GetReturnCode() As MSGReturnCode
    IncomeFactor_GetReturnCode = m_clsIncomeFactor.GetReturnCode
End Function


'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' Function    : IncomeFactor_Initialise
' Description : Delegate to the base class.
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub IncomeFactor_Initialise()
    m_clsIncomeFactor.Initialise
End Sub


'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' Function    : IncomeFactor_SetForm
' Description : Delegate to the base class.
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub IncomeFactor_SetForm(frmParent As Form)
    m_clsIncomeFactor.SetForm frmParent
End Sub


'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' Function    : IncomeFactor_SetIsEdit
' Description : Delegate to the base class.
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub IncomeFactor_SetIsEdit(ByVal bIsEdit As Boolean)
    m_clsIncomeFactor.SetIsEdit bIsEdit
End Sub


'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' Function    : IncomeFactor_SetKeys
' Description : Delegate to the base class.
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub IncomeFactor_SetKeys(colKeys As Collection)
    m_clsIncomeFactor.SetKeys colKeys
End Sub


'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' Function    : IncomeFactor_Terminate
' Description : Delegate to the base class.
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub IncomeFactor_Terminate()
    
    If Not m_clsIncomeFactor Is Nothing Then
        m_clsIncomeFactor.Terminate
    End If
    
End Sub

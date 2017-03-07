VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 3  'UsesTransaction
END
Attribute VB_Name = "IncomeCalcsDO"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

'Workfile:      IncomeCalcsDO.cls
'Copyright:     Copyright © 2002 Marlborough Stirling

'Description:   Income Calculations Data Object which 'supports transactions' only.

'Issues:        Instancing:         MultiUse
'               MTSTransactionMode: UsesTransaction
'-------------------------------------------------------------------------------------------
'BMIDS History:
'
' Prog  Date        Description
'
' MDC   03/07/02    BMIDS00654 Maximum Borrowing Calculations
'-------------------------------------------------------------------------------------------

Implements ObjectControl
Implements IIncomeCalcsDO

Private m_objXMLAssist As XMLAssist
Private m_objErrAssist As ErrAssist
Private m_objContext As ObjectContext

Private Function IIncomeCalcsDO_GetIncomeMultipleSet(ByVal vxmlRequest As IXMLDOMElement) As IXMLDOMNode

On Error GoTo GetDataVbErr

Const strFunctionName As String = "IIncomeCalcsDO_GetIncomeMultipleSet"
    
Dim objIDOAssist As IDOAssist
Dim xmlClassDefDoc As FreeThreadedDOMDocument40
Dim objIomICClassDef As IomICClassDef
    
    Set objIDOAssist = m_objContext.CreateInstance(gstrBASE_COMPONENT & ".DOAssist")
    Set objIomICClassDef = New omICClassDef
    
    Set xmlClassDefDoc = objIomICClassDef.LoadIncomeMultipleSetData()
    Set IIncomeCalcsDO_GetIncomeMultipleSet = objIDOAssist.GetData(vxmlRequest, xmlClassDefDoc)
    
    Set objIomICClassDef = Nothing
    Set xmlClassDefDoc = Nothing
    Set objIDOAssist = Nothing
    
    m_objContext.SetComplete
    
    Exit Function

GetDataVbErr:
    
    Set objIomICClassDef = Nothing
    Set xmlClassDefDoc = Nothing
    Set objIDOAssist = Nothing
    
    m_objErrAssist.AddToErrSource (strFunctionName)
    
    If m_objErrAssist.IsSystemError = True Then
        m_objErrAssist.LogError TypeName(Me), strFunctionName, Err.Number, Err.Description
    End If
    
    If m_objErrAssist.IsSystemError Then
        m_objContext.SetAbort
    Else
        m_objContext.SetComplete
    End If
    
    '   re-raise error for business object to interpret as appropriate
    Err.Raise Err.Number, Err.Source, Err.Description

End Function

Private Sub ObjectControl_Activate()

    Set m_objXMLAssist = New XMLAssist
    Set m_objErrAssist = New ErrAssist
    Set m_objContext = GetObjectContext()

End Sub

Private Function ObjectControl_CanBePooled() As Boolean

    ObjectControl_CanBePooled = True

End Function

Private Sub ObjectControl_Deactivate()

    Set m_objXMLAssist = Nothing
    Set m_objErrAssist = Nothing
    Set m_objContext = Nothing

End Sub
